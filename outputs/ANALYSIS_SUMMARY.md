# Online Retail II — Analysis Summary

## Executive overview

- **Source:** `online_retail_II.xlsx`, sheets: Year 2009-2010, Year 2010-2011
- **Sales period:** 2009-12-01 to 2011-12-09
- **Source rows:** 1,067,371; exact duplicate rows removed: 34,335
- **Included sales lines:** 1,007,913
- **Positive sales revenue:** 20,476,260.45 source price units
- **Orders:** 40,077; average order value: 510.92
- **Identified customers:** 5,878; units sold: 11,205,148
- **Cancellation invoices:** 8,292; negative-quantity return lines: 22,496

Revenue is the sum of quantity × unit price for positive-quantity, positive-price, valid-date lines on invoices that do not start with `C`. Cancellation and return lines are excluded from sales totals and reported separately. Customer analytics include only rows with a customer ID. Revenue is expressed in the workbook's price units; no currency is inferred or converted.

## Highest-revenue countries

| country | revenue | orders |
| --- | --- | --- |
| United Kingdom | 17,410,196.12 | 36,535.00 |
| EIRE | 658,767.31 | 626.00 |
| Netherlands | 554,038.09 | 228.00 |
| Germany | 425,019.71 | 789.00 |
| France | 350,456.09 | 622.00 |

## Highest-revenue products

| stock_code | description | revenue |
| --- | --- | --- |
| M | Manual | 339,226.24 |
| 22423 | REGENCY CAKESTAND 3 TIER | 330,590.32 |
| DOT | DOTCOM POSTAGE | 309,854.11 |
| 85123A | WHITE HANGING HEART T-LIGHT HOLDER | 257,546.20 |
| 23843 | PAPER CRAFT , LITTLE BIRDIE | 168,469.60 |

## RFM customer segments

| segment | customers | revenue | average_customer_revenue | average_orders | average_recency_days |
| --- | --- | --- | --- | --- | --- |
| Champions | 1,971.00 | 13,231,871.08 | 6,713.28 | 13.35 | 29.75 |
| At risk | 968.00 | 2,293,293.77 | 2,369.11 | 6.23 | 289.23 |
| Hibernating | 1,971.00 | 1,016,604.23 | 515.78 | 1.42 | 409.38 |
| Loyal customers | 599.00 | 693,450.88 | 1,157.68 | 2.42 | 35.39 |
| Recent customers | 369.00 | 139,584.30 | 378.28 | 1.00 | 45.38 |

## Analysis coverage

The reproducible pipeline includes workbook profiling, missing-data and duplicate checks, rule-based cleaning, revenue and order KPIs, monthly trends and month-over-month change, geographic and product ranking, weekday/hour patterns, returns, customer-level RFM scoring and segment summaries, and customer cohort retention. Results are also delivered as a formatted Excel workbook with charts, filters, conditional formatting, and an Excel growth formula.

## Caveats

- This is descriptive analysis, not a forecast or causal explanation.
- Cancellation invoices are recognized from invoice IDs beginning with `C`. Negative quantities are separately treated as returns.
- Customer IDs are intentionally excluded from published, customer-level outputs; only aggregated segment summaries are shared.
- Each input row is treated as an item line. Product descriptions and invoice IDs come from the supplied workbook.
- Duplicate removal uses a 64-bit row fingerprint to avoid holding the full transaction table in memory; fingerprint collisions are extraordinarily unlikely but theoretically possible.
