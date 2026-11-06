#!/usr/bin/env python3
"""Reproducible, memory-conscious analysis of the Online Retail II workbook."""

from __future__ import annotations

import argparse
import json
from collections import defaultdict
from itertools import islice
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


SOURCE_COLUMNS = [
    "invoice",
    "stock_code",
    "description",
    "quantity",
    "invoice_date",
    "unit_price",
    "customer_id",
    "country",
]
OUTPUT_DIR = Path("outputs")
CHART_DIR = OUTPUT_DIR / "charts"
BATCH_SIZE = 50_000
PALETTE = {
    "navy": "#17324D",
    "blue": "#2878B5",
    "teal": "#35A7A0",
    "gold": "#F2B134",
    "light": "#F3F6F9",
    "red": "#D9685D",
}


def _merge_grouped(parts: list[pd.DataFrame], keys: list[str], values: list[str]) -> pd.DataFrame:
    if not parts:
        return pd.DataFrame(columns=keys + values)
    combined = pd.concat(parts, ignore_index=True)
    return combined.groupby(keys, as_index=False, dropna=False)[values].sum()


def _add_to_set_map(target: dict[str, set], pairs) -> None:
    for key, value in pairs:
        target[key].add(value)


def _canonical_numeric(value):
    if pd.isna(value):
        return pd.NA
    numeric = float(value)
    if numeric.is_integer():
        return str(int(numeric))
    return format(numeric, ".15g")


def analyze_workbook(source: Path, output_dir: Path = OUTPUT_DIR) -> dict[str, object]:
    if not source.is_file():
        raise FileNotFoundError(f"Workbook not found: {source}")

    chart_dir = output_dir / "charts"
    chart_dir.mkdir(parents=True, exist_ok=True)
    grouped_parts: dict[str, list[pd.DataFrame]] = defaultdict(list)
    unique_fingerprints: set[int] = set()
    all_sales_invoices: set[str] = set()
    month_invoices: dict[str, set[str]] = defaultdict(set)
    month_customers: dict[str, set[str]] = defaultdict(set)
    country_invoices: dict[str, set[str]] = defaultdict(set)
    product_invoices: dict[str, set[str]] = defaultdict(set)
    weekday_invoices: dict[str, set[str]] = defaultdict(set)
    customer_invoices: dict[str, set[str]] = defaultdict(set)
    customer_months: set[tuple[str, str]] = set()
    customer_first_month: dict[str, str] = {}
    customer_last_purchase: dict[str, pd.Timestamp] = {}
    customer_revenue: dict[str, float] = defaultdict(float)
    customer_units: dict[str, int] = defaultdict(int)
    cancellation_invoices: set[str] = set()
    quality = defaultdict(int)
    year_sheets: list[str] = []
    first_sale: pd.Timestamp | None = None
    last_sale: pd.Timestamp | None = None

    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        for sheet_name in workbook.sheetnames:
            year_sheets.append(sheet_name)
            rows = workbook[sheet_name].iter_rows(min_row=2, values_only=True)
            while batch := list(islice(rows, BATCH_SIZE)):
                raw = pd.DataFrame(batch, columns=SOURCE_COLUMNS)
                quality["source_rows"] += len(raw)
                canonical = raw.copy()
                for column in ("invoice", "stock_code", "description", "country"):
                    canonical[column] = canonical[column].astype("string")
                for column in ("quantity", "unit_price", "customer_id"):
                    canonical[column] = pd.to_numeric(canonical[column], errors="coerce").map(
                        _canonical_numeric
                    ).astype("string")
                canonical["invoice_date"] = pd.to_datetime(
                    canonical["invoice_date"], errors="coerce"
                ).map(lambda value: value.isoformat() if pd.notna(value) else pd.NA).astype("string")
                fingerprints = pd.util.hash_pandas_object(canonical, index=False).to_numpy()
                keep = []
                for position, fingerprint in enumerate(fingerprints):
                    fingerprint = int(fingerprint)
                    if fingerprint in unique_fingerprints:
                        quality["exact_duplicate_rows"] += 1
                    else:
                        unique_fingerprints.add(fingerprint)
                        keep.append(position)
                frame = raw.iloc[keep].copy()
                if frame.empty:
                    continue

                frame["invoice"] = frame["invoice"].astype("string").str.strip()
                frame["stock_code"] = frame["stock_code"].astype("string").str.strip()
                frame["description"] = frame["description"].astype("string").str.strip()
                frame["country"] = frame["country"].astype("string").str.strip()
                frame["quantity"] = pd.to_numeric(frame["quantity"], errors="coerce")
                frame["unit_price"] = pd.to_numeric(frame["unit_price"], errors="coerce")
                frame["invoice_date"] = pd.to_datetime(frame["invoice_date"], errors="coerce")
                frame["customer_id"] = pd.to_numeric(
                    frame["customer_id"], errors="coerce"
                ).astype("Int64").astype("string")

                quality["missing_invoice"] += int(frame["invoice"].isna().sum())
                quality["missing_stock_code"] += int(frame["stock_code"].isna().sum())
                quality["missing_description"] += int(frame["description"].isna().sum())
                quality["missing_customer_id"] += int(frame["customer_id"].isna().sum())
                quality["missing_country"] += int(frame["country"].isna().sum())
                quality["invalid_invoice_date"] += int(frame["invoice_date"].isna().sum())
                quality["nonpositive_quantity_rows"] += int(frame["quantity"].le(0).sum())
                quality["nonpositive_price_rows"] += int(frame["unit_price"].le(0).sum())

                is_cancellation = frame["invoice"].str.upper().str.startswith("C", na=False)
                cancellation_invoices.update(frame.loc[is_cancellation, "invoice"].dropna().tolist())
                quality["cancellation_rows"] += int(is_cancellation.sum())
                quality["negative_quantity_return_rows"] += int(frame["quantity"].lt(0).sum())
                quality["zero_quantity_rows"] += int(frame["quantity"].eq(0).sum())
                quality["negative_price_rows"] += int(frame["unit_price"].lt(0).sum())

                frame["country"] = frame["country"].fillna("Unknown")
                cancelled = is_cancellation | frame["quantity"].le(0)
                returns = frame.loc[frame["quantity"].lt(0)].copy()
                if not returns.empty:
                    returns["return_value"] = returns["quantity"].abs() * returns["unit_price"].clip(lower=0)
                    grouped_parts["returns"].append(
                        returns.groupby("country", as_index=False).agg(
                            return_lines=("invoice", "size"),
                            returned_units=("quantity", lambda values: values.abs().sum()),
                            return_value=("return_value", "sum"),
                        )
                    )

                valid_sales = (
                    ~cancelled
                    & frame["quantity"].gt(0)
                    & frame["unit_price"].gt(0)
                    & frame["invoice_date"].notna()
                    & frame["invoice"].notna()
                    & frame["stock_code"].notna()
                )
                sales = frame.loc[valid_sales].copy()
                if sales.empty:
                    continue
                sales["revenue"] = sales["quantity"] * sales["unit_price"]
                sales["month"] = sales["invoice_date"].dt.to_period("M").astype(str)
                sales["weekday"] = sales["invoice_date"].dt.day_name()
                sales["hour"] = sales["invoice_date"].dt.hour
                sales["product"] = sales["stock_code"].fillna("Unknown")

                batch_min = sales["invoice_date"].min()
                batch_max = sales["invoice_date"].max()
                first_sale = batch_min if first_sale is None else min(first_sale, batch_min)
                last_sale = batch_max if last_sale is None else max(last_sale, batch_max)

                grouped_parts["monthly"].append(
                    sales.groupby("month", as_index=False).agg(
                        revenue=("revenue", "sum"),
                        units_sold=("quantity", "sum"),
                        sales_lines=("invoice", "size"),
                    )
                )
                grouped_parts["country"].append(
                    sales.groupby("country", as_index=False).agg(
                        revenue=("revenue", "sum"),
                        units_sold=("quantity", "sum"),
                        sales_lines=("invoice", "size"),
                    )
                )
                grouped_parts["product"].append(
                    sales.groupby(["product", "description"], as_index=False, dropna=False).agg(
                        revenue=("revenue", "sum"),
                        units_sold=("quantity", "sum"),
                        sales_lines=("invoice", "size"),
                    )
                )
                grouped_parts["weekday"].append(
                    sales.groupby("weekday", as_index=False).agg(
                        revenue=("revenue", "sum"),
                        units_sold=("quantity", "sum"),
                        sales_lines=("invoice", "size"),
                    )
                )
                grouped_parts["hour"].append(
                    sales.groupby("hour", as_index=False).agg(
                        revenue=("revenue", "sum"),
                        units_sold=("quantity", "sum"),
                        sales_lines=("invoice", "size"),
                    )
                )

                invoices = sales["invoice"].astype(str)
                month_values = sales["month"].astype(str)
                countries = sales["country"].astype(str)
                products = sales["product"].astype(str)
                all_sales_invoices.update(invoices.unique())
                _add_to_set_map(month_invoices, zip(month_values, invoices))
                _add_to_set_map(country_invoices, zip(countries, invoices))
                _add_to_set_map(product_invoices, zip(products, invoices))
                _add_to_set_map(weekday_invoices, zip(sales["weekday"].astype(str), invoices))

                identified = sales.loc[sales["customer_id"].notna()].copy()
                quality["sales_lines_missing_customer_id"] += len(sales) - len(identified)
                if identified.empty:
                    continue
                identified["customer_id"] = identified["customer_id"].astype(str)
                identified["month"] = identified["month"].astype(str)
                _add_to_set_map(
                    month_customers,
                    zip(identified["month"], identified["customer_id"]),
                )
                customer_months.update(
                    zip(identified["customer_id"], identified["month"])
                )
                for customer, month in zip(identified["customer_id"], identified["month"]):
                    previous = customer_first_month.get(customer)
                    if previous is None or month < previous:
                        customer_first_month[customer] = month

                customer_batch = identified.groupby("customer_id").agg(
                    revenue=("revenue", "sum"),
                    units=("quantity", "sum"),
                    last_purchase=("invoice_date", "max"),
                )
                for customer, record in customer_batch.iterrows():
                    customer_revenue[customer] += float(record["revenue"])
                    customer_units[customer] += int(record["units"])
                    last_purchase = record["last_purchase"]
                    previous = customer_last_purchase.get(customer)
                    if previous is None or last_purchase > previous:
                        customer_last_purchase[customer] = last_purchase
                _add_to_set_map(
                    customer_invoices,
                    zip(
                        identified["customer_id"],
                        identified["invoice"].astype(str),
                    ),
                )
    finally:
        workbook.close()

    if not grouped_parts["monthly"]:
        raise ValueError("No valid positive-price, positive-quantity sales were found.")

    monthly = _merge_grouped(
        grouped_parts["monthly"], ["month"], ["revenue", "units_sold", "sales_lines"]
    ).sort_values("month")
    all_months = pd.period_range(monthly["month"].min(), monthly["month"].max(), freq="M").astype(str)
    monthly = monthly.set_index("month").reindex(all_months, fill_value=0).rename_axis("month").reset_index()
    monthly["orders"] = monthly["month"].map(lambda month: len(month_invoices[month]))
    monthly["customers"] = monthly["month"].map(lambda month: len(month_customers[month]))
    monthly["average_order_value"] = monthly["revenue"].div(monthly["orders"].replace(0, np.nan))
    monthly["mom_revenue_change_pct"] = monthly["revenue"].pct_change() * 100

    country = _merge_grouped(
        grouped_parts["country"], ["country"], ["revenue", "units_sold", "sales_lines"]
    )
    country["orders"] = country["country"].map(lambda item: len(country_invoices[item]))
    country["average_order_value"] = country["revenue"].div(country["orders"].replace(0, np.nan))
    country = country.sort_values("revenue", ascending=False).reset_index(drop=True)

    products = _merge_grouped(
        grouped_parts["product"],
        ["product", "description"],
        ["revenue", "units_sold", "sales_lines"],
    )
    products["orders"] = products["product"].map(lambda item: len(product_invoices[item]))
    products = products.sort_values("revenue", ascending=False).reset_index(drop=True)
    products = products.rename(columns={"product": "stock_code"})

    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday = _merge_grouped(
        grouped_parts["weekday"], ["weekday"], ["revenue", "units_sold", "sales_lines"]
    )
    weekday["orders"] = weekday["weekday"].map(
        lambda day: len(weekday_invoices[day])
    )
    weekday = weekday.set_index("weekday").reindex(weekday_order, fill_value=0).rename_axis("weekday").reset_index()
    weekday["average_line_revenue"] = weekday["revenue"].div(weekday["sales_lines"].replace(0, np.nan))

    hour = _merge_grouped(
        grouped_parts["hour"], ["hour"], ["revenue", "units_sold", "sales_lines"]
    ).sort_values("hour")
    hour["average_line_revenue"] = hour["revenue"].div(hour["sales_lines"].replace(0, np.nan))

    returns = _merge_grouped(
        grouped_parts["returns"], ["country"], ["return_lines", "returned_units", "return_value"]
    )
    returns = returns.sort_values("return_value", ascending=False).reset_index(drop=True)

    recency_reference = last_sale + pd.Timedelta(days=1)
    rfm = pd.DataFrame(
        {
            "customer_id": list(customer_revenue),
            "monetary_value": list(customer_revenue.values()),
            "units_purchased": [customer_units[c] for c in customer_revenue],
            "orders": [len(customer_invoices[c]) for c in customer_revenue],
            "last_purchase": [customer_last_purchase[c] for c in customer_revenue],
        }
    )
    rfm["recency_days"] = (recency_reference - rfm["last_purchase"]).dt.days
    customer_total = len(rfm)
    for column, labels, ascending in (
        ("recency_days", [4, 3, 2, 1], True),
        ("orders", [1, 2, 3, 4], False),
        ("monetary_value", [1, 2, 3, 4], False),
    ):
        rank = rfm[column].rank(method="first", ascending=ascending)
        quartile = np.ceil(rank / customer_total * 4).clip(1, 4).astype(int)
        rfm[f"{column}_score"] = 5 - quartile
    rfm["rfm_score"] = (
        rfm["recency_days_score"].astype(str)
        + rfm["orders_score"].astype(str)
        + rfm["monetary_value_score"].astype(str)
    )
    rfm["segment"] = np.select(
        [
            rfm["recency_days_score"].ge(3) & rfm["orders_score"].ge(3),
            rfm["recency_days_score"].ge(3) & rfm["orders_score"].ge(2),
            rfm["recency_days_score"].ge(3),
            rfm["orders_score"].ge(3),
        ],
        ["Champions", "Loyal customers", "Recent customers", "At risk"],
        default="Hibernating",
    )
    customer_segments = (
        rfm.groupby("segment", as_index=False)
        .agg(
            customers=("customer_id", "size"),
            revenue=("monetary_value", "sum"),
            average_customer_revenue=("monetary_value", "mean"),
            average_orders=("orders", "mean"),
            average_recency_days=("recency_days", "mean"),
        )
        .sort_values("revenue", ascending=False)
        .reset_index(drop=True)
    )

    cohort_first = customer_first_month
    cohort_counts: dict[tuple[str, int], set[str]] = defaultdict(set)
    cohort_sizes: dict[str, set[str]] = defaultdict(set)
    for customer, transaction_month in customer_months:
        cohort = cohort_first[customer]
        cohort_sizes[cohort].add(customer)
        cohort_period = pd.Period(cohort, freq="M")
        transaction_period = pd.Period(transaction_month, freq="M")
        age = (transaction_period.year - cohort_period.year) * 12 + transaction_period.month - cohort_period.month
        if age >= 0:
            cohort_counts[(cohort, age)].add(customer)
    cohort_rows = []
    for cohort in sorted(cohort_sizes):
        size = len(cohort_sizes[cohort])
        for (row_cohort, age), customers in sorted(cohort_counts.items()):
            if row_cohort == cohort:
                cohort_rows.append(
                    {
                        "cohort_month": cohort,
                        "months_since_first_purchase": age,
                        "cohort_customers": size,
                        "retained_customers": len(customers),
                        "retention_pct": 100 * len(customers) / size,
                    }
                )
    cohort_long = pd.DataFrame(cohort_rows)
    cohort_matrix = cohort_long.pivot(
        index="cohort_month",
        columns="months_since_first_purchase",
        values="retention_pct",
    )

    sales_revenue = float(monthly["revenue"].sum())
    order_count = len(all_sales_invoices)
    customer_count = len(customer_revenue)
    metrics: dict[str, object] = {
        "source_file": source.name,
        "source_rows": int(quality["source_rows"]),
        "exact_duplicate_rows": int(quality["exact_duplicate_rows"]),
        "analyzed_sales_lines": int(monthly["sales_lines"].sum()),
        "excluded_non_sales_rows": int(
            quality["source_rows"]
            - quality["exact_duplicate_rows"]
            - int(monthly["sales_lines"].sum())
        ),
        "unique_orders": order_count,
        "identified_customers": customer_count,
        "unidentified_sales_lines": int(quality["sales_lines_missing_customer_id"]),
        "positive_sales_revenue": sales_revenue,
        "average_order_value": sales_revenue / order_count if order_count else 0,
        "units_sold": int(monthly["units_sold"].sum()),
        "sales_date_start": first_sale.strftime("%Y-%m-%d"),
        "sales_date_end": last_sale.strftime("%Y-%m-%d"),
        "source_sheets": year_sheets,
        "unique_cancellation_invoices": len(cancellation_invoices),
        "cancellation_rows": int(quality["cancellation_rows"]),
        "negative_quantity_return_rows": int(quality["negative_quantity_return_rows"]),
        "negative_quantity_return_value": float(returns["return_value"].sum()) if not returns.empty else 0,
        "missing_customer_id_rows": int(quality["missing_customer_id"]),
    }

    # Keep data-quality counts explicit and reproducible alongside the headline KPIs.
    quality_rows = [
        ("Source rows", quality["source_rows"], "Rows across both worksheets"),
        ("Exact duplicate rows", quality["exact_duplicate_rows"], "Repeated source rows removed before analysis"),
        ("Missing invoice", quality["missing_invoice"], "Invoice identifier is blank"),
        ("Missing stock code", quality["missing_stock_code"], "Stock code is blank"),
        ("Missing description", quality["missing_description"], "Description is blank"),
        ("Missing customer ID", quality["missing_customer_id"], "Customer-level analysis cannot identify these rows"),
        ("Sales lines without customer ID", quality["sales_lines_missing_customer_id"], "Included in sales totals but not customer analytics"),
        ("Missing country", quality["missing_country"], "Country is blank; sales analysis groups it as Unknown"),
        ("Invalid invoice date", quality["invalid_invoice_date"], "Invoice date could not be parsed"),
        ("Cancellation rows", quality["cancellation_rows"], "Invoice identifier starts with C"),
        ("Negative-quantity return rows", quality["negative_quantity_return_rows"], "Negative quantity; excluded from sales and RFM"),
        ("Zero-quantity rows", quality["zero_quantity_rows"], "Zero quantity; excluded from sales and RFM"),
        ("Nonpositive-quantity rows", quality["nonpositive_quantity_rows"], "Quantity is zero or negative"),
        ("Nonpositive-price rows", quality["nonpositive_price_rows"], "Unit price is zero or negative"),
    ]
    data_quality = pd.DataFrame(quality_rows, columns=["check", "rows", "interpretation"])

    summary = pd.DataFrame(
        [
            ("Source rows", metrics["source_rows"], "Rows in both source worksheets"),
            ("Exact duplicate rows removed", metrics["exact_duplicate_rows"], "Duplicate source rows"),
            ("Sales lines analyzed", metrics["analyzed_sales_lines"], "Positive quantity and price, valid date, non-cancellation"),
            ("Orders", order_count, "Distinct invoices from included sales lines"),
            ("Identified customers", customer_count, "Distinct customers represented in RFM analysis"),
            ("Revenue", sales_revenue, "Positive-sale line value; returns/cancellations excluded"),
            ("Average order value", metrics["average_order_value"], "Included positive sales revenue divided by distinct invoices"),
            ("Units sold", metrics["units_sold"], "Sum of positive quantities in included sales lines"),
            ("Cancellation invoices", metrics["unique_cancellation_invoices"], "Distinct invoices with an ID starting with C"),
            ("Negative-quantity return lines", metrics["negative_quantity_return_rows"], "Negative quantity rows reported separately"),
        ],
        columns=["metric", "value", "definition"],
    )
    cohort_summary = cohort_long.copy()

    output_dir.mkdir(parents=True, exist_ok=True)
    exports = {
        "monthly_sales.csv": monthly,
        "country_analysis.csv": country,
        "product_analysis.csv": products,
        "weekday_analysis.csv": weekday,
        "hour_analysis.csv": hour,
        "customer_segments.csv": customer_segments,
        "cohort_retention.csv": cohort_summary,
        "cohort_retention_matrix.csv": cohort_matrix.reset_index(),
        "returns_analysis.csv": returns,
        "data_quality.csv": data_quality,
        "kpi_summary.csv": summary,
    }
    for filename, data in exports.items():
        data.to_csv(output_dir / filename, index=False, encoding="utf-8-sig")
    (output_dir / "metrics.json").write_text(
        json.dumps(metrics, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    create_charts(monthly, country, products, weekday, cohort_matrix, chart_dir)
    create_excel_workbook(
        output_dir / "online_retail_ii_analysis.xlsx",
        metrics,
        summary,
        data_quality,
        monthly,
        country,
        products.head(100),
        weekday,
        hour,
        customer_segments,
        cohort_matrix,
        returns,
    )
    write_summary(output_dir / "ANALYSIS_SUMMARY.md", metrics, country, products, customer_segments)
    return metrics


def create_charts(monthly, country, products, weekday, cohort_matrix, output_dir: Path) -> None:
    sns.set_theme(style="whitegrid", palette=[PALETTE["blue"], PALETTE["teal"], PALETTE["gold"]])
    fig, ax = plt.subplots(figsize=(12, 5))
    sns.lineplot(data=monthly, x="month", y="revenue", marker="o", color=PALETTE["blue"], ax=ax)
    ax.set(title="Monthly sales revenue", xlabel="Invoice month", ylabel="Revenue (dataset price units)")
    ax.tick_params(axis="x", rotation=45)
    fig.tight_layout()
    fig.savefig(output_dir / "monthly_revenue.png", dpi=180)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 6))
    top = country.head(10).sort_values("revenue")
    sns.barplot(data=top, x="revenue", y="country", color=PALETTE["teal"], ax=ax)
    ax.set(title="Top 10 countries by sales revenue", xlabel="Revenue (dataset price units)", ylabel="")
    fig.tight_layout()
    fig.savefig(output_dir / "top_countries.png", dpi=180)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(11, 7))
    top = products.head(15).sort_values("revenue")
    sns.barplot(data=top, x="revenue", y="description", color=PALETTE["blue"], ax=ax)
    ax.set(title="Top 15 products by sales revenue", xlabel="Revenue (dataset price units)", ylabel="")
    fig.tight_layout()
    fig.savefig(output_dir / "top_products.png", dpi=180)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(10, 5))
    sns.barplot(data=weekday, x="weekday", y="revenue", color=PALETTE["gold"], ax=ax)
    ax.set(title="Sales revenue by weekday", xlabel="", ylabel="Revenue (dataset price units)")
    ax.tick_params(axis="x", rotation=30)
    fig.tight_layout()
    fig.savefig(output_dir / "weekday_revenue.png", dpi=180)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(13, 7))
    sns.heatmap(cohort_matrix, cmap="YlGnBu", annot=True, fmt=".0f", vmin=0, vmax=100, ax=ax)
    ax.set(title="Customer cohort retention (%)", xlabel="Months since first purchase", ylabel="First purchase cohort")
    fig.tight_layout()
    fig.savefig(output_dir / "cohort_retention.png", dpi=180)
    plt.close(fig)


def _write_sheet(workbook, name: str, data: pd.DataFrame, percentage_columns=()) -> None:
    sheet = workbook.create_sheet(name)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A2"
    sheet.append([str(column) for column in data.columns])
    for row in data.itertuples(index=False, name=None):
        sheet.append([None if pd.isna(value) else value.item() if isinstance(value, np.generic) else value for value in row])

    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = PatternFill("solid", fgColor="17324D")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.row_dimensions[1].height = 30

    for column_cells in sheet.columns:
        values = list(column_cells)
        header = str(values[0].value or "")
        width = min(max(max((len(str(cell.value or "")) for cell in values[:150]), default=0) + 2, 12), 42)
        sheet.column_dimensions[values[0].column_letter].width = width
        if "revenue" in header.lower() or "value" in header.lower() or "price" in header.lower():
            for cell in values[1:]:
                cell.number_format = '#,##0.00;[Red]-#,##0.00;0'
        if "pct" in header.lower() or header in percentage_columns:
            for cell in values[1:]:
                cell.number_format = '0.0"%"'
    if sheet.max_row > 1 and sheet.max_column:
        table_name = "T" + "".join(character for character in name.title() if character.isalnum())
        table = Table(displayName=table_name, ref=sheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        sheet.add_table(table)
    return sheet


def create_excel_workbook(
    path: Path,
    metrics: dict,
    summary: pd.DataFrame,
    data_quality: pd.DataFrame,
    monthly: pd.DataFrame,
    country: pd.DataFrame,
    products: pd.DataFrame,
    weekday: pd.DataFrame,
    hour: pd.DataFrame,
    customer_segments: pd.DataFrame,
    cohort_matrix: pd.DataFrame,
    returns: pd.DataFrame,
) -> None:
    workbook = Workbook()
    overview = workbook.active
    overview.title = "Overview"
    overview.sheet_view.showGridLines = False
    overview["A1"] = "ONLINE RETAIL II | SALES & CUSTOMER ANALYTICS"
    overview["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    overview["A1"].fill = PatternFill("solid", fgColor="17324D")
    overview.merge_cells("A1:F1")
    overview["A2"] = (
        f"Source: {metrics['source_file']} | Sales period: "
        f"{metrics['sales_date_start']} to {metrics['sales_date_end']}"
    )
    overview.merge_cells("A2:F2")
    overview["A2"].font = Font(italic=True, color="536579")
    for cell in ("A4", "C4", "E4"):
        overview[cell].font = Font(bold=True, color="FFFFFF")
        overview[cell].fill = PatternFill("solid", fgColor="2878B5")
    kpis = [
        ("A4", "Sales revenue", "A5", float(metrics["positive_sales_revenue"]), '#,##0.00'),
        ("C4", "Orders", "C5", int(metrics["unique_orders"]), '#,##0'),
        ("E4", "Identified customers", "E5", int(metrics["identified_customers"]), '#,##0'),
        ("A7", "Average order value", "A8", float(metrics["average_order_value"]), '#,##0.00'),
        ("C7", "Units sold", "C8", int(metrics["units_sold"]), '#,##0'),
        ("E7", "Cancellation invoices", "E8", int(metrics["unique_cancellation_invoices"]), '#,##0'),
    ]
    for label_cell, label, value_cell, value, number_format in kpis:
        overview[label_cell] = label
        overview[label_cell].fill = PatternFill("solid", fgColor="2878B5")
        overview[label_cell].font = Font(bold=True, color="FFFFFF")
        overview[value_cell] = value
        overview[value_cell].font = Font(bold=True, size=15, color="17324D")
        overview[value_cell].number_format = number_format
    overview["A10"] = "How to read this workbook"
    overview["A10"].font = Font(bold=True, size=12, color="17324D")
    overview["A11"] = "Sales exclude cancellation invoices, negative/zero quantities, nonpositive prices, invalid dates, and exact duplicate rows."
    overview.merge_cells("A11:F11")
    overview["A12"] = "Returns are summarized separately. Revenue is quantity × unit price in the source workbook's price units; no currency conversion is assumed."
    overview.merge_cells("A12:F12")
    overview["A13"] = "RFM and cohort analysis use identified customers only. Customer IDs are not included in published outputs."
    overview.merge_cells("A13:F13")
    overview["A14"] = "The original transaction-level workbook is deliberately not included; use the documented input path to regenerate results."
    overview.merge_cells("A14:F14")
    for row in range(11, 15):
        overview[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="top")
        overview.row_dimensions[row].height = 32

    sheet_summary = _write_sheet(workbook, "KPI Summary", summary)
    quality_sheet = _write_sheet(workbook, "Data Quality", data_quality)
    monthly_sheet = _write_sheet(workbook, "Monthly Sales", monthly)
    _write_sheet(workbook, "Country Analysis", country)
    _write_sheet(workbook, "Product Analysis", products)
    _write_sheet(workbook, "Weekday Analysis", weekday)
    _write_sheet(workbook, "Hour Analysis", hour)
    _write_sheet(workbook, "Customer Segments", customer_segments)
    cohort_sheet = _write_sheet(workbook, "Cohort Retention", cohort_matrix.reset_index())
    _write_sheet(workbook, "Returns Analysis", returns)
    dictionary = pd.DataFrame(
        [
            ("Revenue", "Quantity × unit price for included positive-sale lines; source price units retained."),
            ("Order", "Distinct invoice ID among included positive-sale lines."),
            ("Cancellation", "Invoice ID starts with C; reported separately and excluded from sales."),
            ("Return", "Negative-quantity transaction line; summarized separately and excluded from sales/RFM."),
            ("RFM recency", "Days from the most recent included sale to one day after the last sale in the dataset."),
            ("Cohort retention", "Share of customers in a first-purchase month who also purchased in each later month."),
            ("Missing customer ID", "Sales remain in revenue totals but cannot contribute to customer-level metrics."),
        ],
        columns=["term", "definition"],
    )
    _write_sheet(workbook, "Data Dictionary", dictionary)
    for sheet in (monthly_sheet, quality_sheet, cohort_sheet, sheet_summary):
        sheet.auto_filter.ref = sheet.dimensions
    if monthly_sheet.max_row > 3:
        formula_column = monthly_sheet.max_column + 1
        monthly_sheet.cell(1, formula_column, "Excel MoM Change %")
        monthly_sheet.cell(1, formula_column).font = Font(color="FFFFFF", bold=True)
        monthly_sheet.cell(1, formula_column).fill = PatternFill("solid", fgColor="17324D")
        for row in range(2, monthly_sheet.max_row + 1):
            monthly_sheet.cell(
                row,
                formula_column,
                None if row == 2 else f'=IFERROR(B{row}/B{row - 1}-1,"")',
            )
            monthly_sheet.cell(row, formula_column).number_format = "0.0%"
        next(iter(monthly_sheet.tables.values())).ref = (
            f"A1:{get_column_letter(formula_column)}{monthly_sheet.max_row}"
        )
        monthly_sheet.column_dimensions[get_column_letter(formula_column)].width = 23
    if country.shape[0] > 0:
        country_sheet = workbook["Country Analysis"]
        country_sheet.conditional_formatting.add(
            f"B2:B{country_sheet.max_row}",
            DataBarRule(start_type="min", end_type="max", color="35A7A0"),
        )
    if len(cohort_matrix.columns) > 0 and cohort_sheet.max_row > 1:
        cohort_sheet.conditional_formatting.add(
            f"B2:{cohort_sheet.cell(cohort_sheet.max_row, cohort_sheet.max_column).coordinate}",
            ColorScaleRule(
                start_type="num",
                start_value=0,
                start_color="F8696B",
                mid_type="num",
                mid_value=50,
                mid_color="FFEB84",
                end_type="num",
                end_value=100,
                end_color="63BE7B",
            ),
        )
    if monthly_sheet.max_row > 2:
        chart = LineChart()
        chart.title = "Monthly sales revenue"
        chart.y_axis.title = "Revenue"
        chart.x_axis.title = "Month"
        chart.add_data(Reference(monthly_sheet, min_col=2, min_row=1, max_row=monthly_sheet.max_row), titles_from_data=True)
        chart.set_categories(Reference(monthly_sheet, min_col=1, min_row=2, max_row=monthly_sheet.max_row))
        overview.add_chart(chart, "A16")
    if len(country) > 0:
        country_sheet = workbook["Country Analysis"]
        chart = BarChart()
        chart.type = "bar"
        chart.title = "Top countries by revenue"
        chart.add_data(Reference(country_sheet, min_col=2, min_row=1, max_row=min(country_sheet.max_row, 11)), titles_from_data=True)
        chart.set_categories(Reference(country_sheet, min_col=1, min_row=2, max_row=min(country_sheet.max_row, 11)))
        overview.add_chart(chart, "G4")
    overview.column_dimensions["A"].width = 28
    overview.column_dimensions["B"].width = 18
    overview.column_dimensions["C"].width = 22
    overview.column_dimensions["D"].width = 4
    overview.column_dimensions["E"].width = 24
    overview.column_dimensions["F"].width = 18
    for row in overview.iter_rows(min_row=1, max_row=14, min_col=1, max_col=6):
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.save(path)


def write_summary(path: Path, metrics: dict, country: pd.DataFrame, products: pd.DataFrame, segments: pd.DataFrame) -> None:
    def markdown_table(data: pd.DataFrame) -> str:
        headers = [str(column) for column in data.columns]
        rows = [[f"{value:,.2f}" if isinstance(value, (int, float, np.number)) else str(value) for value in row] for row in data.itertuples(index=False, name=None)]
        return "\n".join(
            ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
            + ["| " + " | ".join(row) + " |" for row in rows]
        )

    top_countries = markdown_table(country.head(5)[["country", "revenue", "orders"]])
    top_products = markdown_table(products.head(5)[["stock_code", "description", "revenue"]])
    segment_table = markdown_table(segments)
    text = f"""# Online Retail II — Analysis Summary

## Executive overview

- **Source:** `{metrics['source_file']}`, sheets: {', '.join(metrics['source_sheets'])}
- **Sales period:** {metrics['sales_date_start']} to {metrics['sales_date_end']}
- **Source rows:** {metrics['source_rows']:,}; exact duplicate rows removed: {metrics['exact_duplicate_rows']:,}
- **Included sales lines:** {metrics['analyzed_sales_lines']:,}
- **Positive sales revenue:** {metrics['positive_sales_revenue']:,.2f} source price units
- **Orders:** {metrics['unique_orders']:,}; average order value: {metrics['average_order_value']:,.2f}
- **Identified customers:** {metrics['identified_customers']:,}; units sold: {metrics['units_sold']:,}
- **Cancellation invoices:** {metrics['unique_cancellation_invoices']:,}; negative-quantity return lines: {metrics['negative_quantity_return_rows']:,}

Revenue is the sum of quantity × unit price for positive-quantity, positive-price, valid-date lines on invoices that do not start with `C`. Cancellation and return lines are excluded from sales totals and reported separately. Customer analytics include only rows with a customer ID. Revenue is expressed in the workbook's price units; no currency is inferred or converted.

## Highest-revenue countries

{top_countries}

## Highest-revenue products

{top_products}

## RFM customer segments

{segment_table}

## Analysis coverage

The reproducible pipeline includes workbook profiling, missing-data and duplicate checks, rule-based cleaning, revenue and order KPIs, monthly trends and month-over-month change, geographic and product ranking, weekday/hour patterns, returns, customer-level RFM scoring and segment summaries, and customer cohort retention. Results are also delivered as a formatted Excel workbook with charts, filters, conditional formatting, and an Excel growth formula.

## Caveats

- This is descriptive analysis, not a forecast or causal explanation.
- Cancellation invoices are recognized from invoice IDs beginning with `C`. Negative quantities are separately treated as returns.
- Customer IDs are intentionally excluded from published, customer-level outputs; only aggregated segment summaries are shared.
- Each input row is treated as an item line. Product descriptions and invoice IDs come from the supplied workbook.
- Duplicate removal uses a 64-bit row fingerprint to avoid holding the full transaction table in memory; fingerprint collisions are extraordinarily unlikely but theoretically possible.
"""
    path.write_text(text, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, type=Path, help="Path to the locally supplied .xlsx workbook")
    parser.add_argument("--output", default=OUTPUT_DIR, type=Path, help="Directory for generated analytics")
    arguments = parser.parse_args()
    metrics = analyze_workbook(arguments.input, arguments.output)
    print(
        "Analysis complete: "
        f"{metrics['source_rows']:,} source rows, "
        f"{metrics['unique_orders']:,} orders, "
        f"{metrics['identified_customers']:,} identified customers, "
        f"revenue {metrics['positive_sales_revenue']:,.2f}."
    )
    print(f"Outputs: {arguments.output.resolve()}")


if __name__ == "__main__":
    main()
