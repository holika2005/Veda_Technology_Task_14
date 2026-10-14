# Task 14 — Basic Sales Summary

## Objective
Calculate total sales, average sales, and transaction count, then summarize three KPIs.

## Tool
Microsoft Excel.

## Dataset note
This package uses a **Superstore-style practice dataset** with common order, region, category, quantity, sales, and profit fields. The official Sample Superstore file was not attached with the task screenshot, so this is not represented as the official dataset.

## Files
- `Task_14_Basic_Sales_Summary.xlsx`: `Orders` and `Summary` sheets.
- `superstore_style_sales_data.csv`: source data export.
- `sales_kpi_summary.csv`: KPI summary export.
- `README.md`: project documentation.

## Excel formulas
- Total Sales: `=SUM(Orders[Sales])`
- Average Sales: `=AVERAGE(Orders[Sales])`
- Transaction Count: `=COUNTA(Orders[Order ID])`

## Results for this practice dataset
- Transactions: 20
- Total sales: $3,075.19
- Average sales per transaction row: $153.76

## How to use
1. Open the workbook in Excel.
2. Review the `Orders` sheet and verify sales values.
3. Open `Summary` to review KPI formulas and calculated values.
4. Replace practice rows with the official dataset if supplied, keeping the table name `Orders`.

## Interview questions
**What is a KPI?** A Key Performance Indicator is a measurable value used to track performance against a business objective.

**How is average different from total?** Total adds all transaction sales; average divides the total by the number of transaction rows.

## Data note
This is a practice dataset, not the official Sample Superstore source. No claims are made that it is the original dataset.
