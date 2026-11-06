# Online Retail II Analytics

End-to-end, reproducible Excel and Python analytics for the **Online Retail II** workbook. This standalone project profiles and cleans both annual worksheets, analyzes sales, customers, products, returns, and seasonality, and generates a stakeholder-ready Excel workbook, charts, and aggregated exports.

## Analysis included

- **Workbook and data quality:** sheet and row counts, types, missing values, exact duplicate detection, invalid dates, nonpositive quantities/prices, cancellations, and negative-quantity returns.
- **Preparation:** consistent column types, trimmed labels, deduplication, explicit positive-sales rules, and documented handling of missing customer IDs.
- **Descriptive and financial KPIs:** revenue, units, invoice/order counts, average order value, and sales-line counts.
- **Sales trends:** monthly revenue, volume, orders, unique buyers, and month-over-month revenue change.
- **Products and markets:** product and country revenue rankings, order counts, units, and average order value by country.
- **Customer analytics:** RFM (recency, frequency, monetary value) quartile scores and aggregated customer segments.
- **Retention:** first-purchase-month cohort analysis and customer retention rates over subsequent months.
- **Time patterns:** sales by weekday and hour.
- **Returns:** negative-quantity return line, unit, and value summaries, separate from positive sales.
- **Presentation:** Excel dashboard and charts, sortable/filterable tables, data bars, a retention heatmap, and five exported charts.

## Source data

Obtain the **Online Retail II** Excel file from the [UCI Machine Learning Repository](https://archive.ics.uci.edu/dataset/502/online+retail+ii). The original workbook is approximately 44 MB and is intentionally not committed here: it is transaction-level data, unnecessary for distribution, and can be supplied locally when running the analysis.

Save it as `data/online_retail_II.xlsx` (the `data/` directory is ignored by Git), or pass its path explicitly. Do not publish transaction-level customer identifiers. Generated outputs contain aggregated customer segments and cohort data only.

## Run

Python 3.10+ is required. In PowerShell from the repository root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python analytics.py --input "data\online_retail_II.xlsx"
```

To use the supplied local file at another location:

```powershell
python analytics.py --input "C:\path\to\online_retail_II.xlsx"
```

The script reads worksheet rows in bounded batches to avoid loading the million-row workbook into memory at once. Results are regenerated in `outputs/`; existing results with the same filenames are replaced.

The notebook can be run in VS Code with the **Jupyter** extension and the same `.venv` Python kernel.

## Deliverables

| Output | Contents |
| --- | --- |
| `outputs/online_retail_ii_analysis.xlsx` | Dashboard, Excel charts, KPI and quality tables, monthly sales, country/product results, weekday/hour profiles, customer segments, cohort retention, returns, and data dictionary |
| `outputs/ANALYSIS_SUMMARY.md` | Automatically generated executive summary and findings |
| `outputs/*.csv` | Aggregated analysis and quality tables for reuse in Excel, Python, or BI tools |
| `outputs/metrics.json` | Machine-readable headline metrics and source profile |
| `outputs/charts/*.png` | Monthly revenue, top countries, top products, weekday revenue, and cohort retention |
| `notebooks/online_retail_ii_analysis.ipynb` | Guided, rerunnable analysis and output review |

## Interpretation and definitions

- **Sales revenue** is `quantity × unit price` for valid-date rows with positive quantity and price whose invoice ID does not begin with `C`. Cancellation and negative-quantity rows are excluded from sales KPIs.
- **Returns** are negative-quantity rows. They are reported separately at absolute unit/value amounts; the source does not establish that each return has a matching original sale.
- **Orders** are distinct included invoice IDs. **Average order value** is included sales revenue divided by distinct included invoices.
- **Customer analysis** requires an identified customer ID. Revenue lines without an ID still count in product, country, and total sales results.
- **RFM** recency is measured as days since the most recent identified sale, frequency is distinct included invoices, and monetary value is included sales revenue. Scores are quartile-based within this dataset; segments are descriptive, not predictions.
- **Cohort retention** counts identified customers from their first purchase month who make a purchase in each later month. It is not a causal or lifetime-value estimate.
- Revenue stays in the source workbook's price units. No currency symbol, tax, shipping, or currency conversion is inferred.
- Exact duplicate rows are identified with a 64-bit row fingerprint. A collision is theoretically possible but extraordinarily unlikely.

## Guided workbook review

Open `outputs/online_retail_ii_analysis.xlsx` in Excel or a compatible spreadsheet application. Start at **Overview**, filter or sort the formatted output tables, inspect the **Monthly Sales** growth formula, compare the **Cohort Retention** heatmap, and check **Data Quality** and **Data Dictionary** before interpreting the rankings. The `outputs/` directory is generated and is not required in Git; run the pipeline to create it locally.
