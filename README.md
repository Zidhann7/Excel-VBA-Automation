# Excel VBA Automation – Sales Data Formatting, PivotTables & Dashboard

An Excel 2024 VBA automation project that turns a raw sales dataset into a structured, formatted analysis workbook with automatically generated PivotTables and a refreshable sales dashboard.

The project is designed to reduce repetitive Excel work: format the source data, build summary analyses, and generate a dashboard from the summaries using VBA macros and buttons.

---

## Project Overview

This project uses a sales dataset containing **1,000 transaction records and 24 columns**. It is organised into three main worksheets:

| Worksheet | Purpose |
|---|---|
| `Dataset` | Original sales records and source fields. |
| `Pivot` | Summary analysis generated with Excel PivotTables. |
| `Dashboard` | KPI cards, charts, and a Top 5 orders section. |

The intended workflow is:

```text
Dataset
   │
   ▼
Format Sales Data (VBA)
   │
   ▼
Create / Refresh PivotTables (VBA)
   │
   ▼
Build / Refresh Dashboard (VBA)
   │
   ▼
Sales Analysis Dashboard
```

## Main Features

### 1. Automated Dataset Formatting

The formatting macro is designed for the worksheet named `Dataset`.

- Detects the used dataset range automatically.
- Creates or reuses the structured Excel Table `tblSalesData`.
- Applies a simple purple table style with filter dropdowns and alternating row shading.
- Applies consistent font, alignment, row heights, and column widths.
- Applies Indian Rupee (`₹`) display formats to monetary fields.
- Applies percentage display formats to discount and profit-margin fields.
- Applies date display formats to date cells that contain real Excel date values.
- Includes a helper to remove unwanted Freeze Panes or Split settings.

**Data integrity:** The formatting macro is intended to change presentation only. It does not trim or rewrite text, convert text dates, sort records, or delete source records and columns.

### 2. Automated PivotTable Generation

The PivotTable macro creates or refreshes a separate `Pivot` worksheet using the source data. The planned PivotTable analyses are:

| PivotTable | Analysis |
|---|---|
| `PT_Segment` | Sales by customer segment. |
| `PT_Trend` | Sales and profit trend by year/month. |
| `PT_Monthly` | Monthly sales performance by year. |
| `PT_Category` | Sales by product category. |
| `PT_Region` | Sales and profit by region. |
| `PT_Discount` | Profit and discount analysis. |
| `PT_Quantity` | Quantity by customer segment. |
| `PT_Top5` | Top 5 orders ranked by sales. |

The PivotTable layout uses separate column blocks with spacing between summaries. A staging/source copy may be used for date grouping so the original `Dataset` values are not rewritten.

### 3. Automated Sales Dashboard

The dashboard macro creates or rebuilds the `Dashboard` worksheet using the PivotTable summaries.

The dashboard is designed to include six KPI cards:

- Total Sales
- Total Profit
- Total Orders
- Total Customers
- Total Quantity Sold
- Profit Margin

It also includes six visual analyses:

- Sales by Customer Segment
- Sales and Profit Trend
- Monthly Sales Performance
- Profit vs. Discount
- Sales by Category
- Sales by Region

A Top 5 Orders section highlights the highest-sales orders. The layout uses a purple theme, a title banner, KPI cards, and a chart grid. A refresh button is provided to rebuild the dashboard after the source or PivotTable data is updated.

> **Current scope note:** Slicer-based interactivity and slicer-responsive KPI calculations are not included in the described implementation. The dashboard is refreshable through VBA; do not describe it as slicer-driven unless that functionality is added and tested.

---

## Dataset Fields

The source dataset contains these 24 fields:

```text
Order_ID
Order_Date
Customer_ID
Customer_Name
Customer_Segment
Region
State
City
Sales_Channel
Product_ID
Product_Name
Category
Sub_Category
Quantity
Unit_Price_INR
Discount_Pct
Sales_Amount_INR
Cost_Amount_INR
Profit_INR
Profit_Margin_Pct
Payment_Method
Order_Status
Ship_Date
Delivery_Days
```

## VBA Macro Entry Points

The project is organised around separate VBA modules. Exact procedure names can vary depending on the version of the module imported into the workbook; the main entry points used in this project are:

| Task | Macro / action |
|---|---|
| Format the dataset | `CleanAndFormatData` |
| Remove split/frozen panes | `FixDatasetSplitView` (or `FixDatasetSplitViewOnly` in the standalone fix module) |
| Create or refresh PivotTables | `CreatePivotSheets` |
| Build the dashboard for the first time | `BuildSalesDashboard` (or `BuildDashboard`, depending on the dashboard module version) |
| Refresh the dashboard | `RefreshDashboard` |

Use the macro names shown in the imported VBA module if they differ from this table.

---

## Requirements

- Microsoft Excel **2024 desktop** for Windows.
- A workbook containing the `Dataset` worksheet and the expected 24 headers.
- Macros enabled for the workbook.
- Save the workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**.

Excel for the web does not run VBA macros. Keep a backup copy of the workbook before running macros that rebuild PivotTables or dashboard objects.

## Setup and Usage

1. Open the workbook in Excel 2024 desktop.
2. Save a working copy as `Excel_VBA_Automation.xlsm` (or another `.xlsm` filename).
3. Press `Alt + F11` to open the Visual Basic Editor.
4. In the editor, choose **File → Import File** and import the required `.bas` modules.
5. Return to Excel and run the dataset formatting macro `CleanAndFormatData`.
6. Run the PivotTable macro `CreatePivotSheets` (or the button-creation macro included in the PivotTable module).
7. Run the dashboard build macro once to create the `Dashboard` worksheet and its refresh button.
8. After updating the source records, run the PivotTable refresh/build step and then click **Refresh Dashboard** to update the dashboard.

### Recommended Run Order

```text
1. CleanAndFormatData
2. CreatePivotSheets
3. BuildSalesDashboard
4. RefreshDashboard  (after later data updates)
```

The dashboard should be refreshed after the PivotTables are refreshed, so its charts and KPI summaries use the latest PivotTable results.

---

## Data Integrity and Implementation Notes

- The `Dataset` worksheet is the source of truth.
- The formatting workflow is intended to preserve the original source values and record order.
- Date fields in the supplied file may be stored as ISO-style text (for example, `2025-09-17`). Applying an Excel number format does not convert text into a true Excel date. Date grouping therefore needs valid date values in a staging copy or an explicit conversion step in that staging area.
- Currency and percentage number formats affect how values are displayed; they do not calculate or change the underlying values.
- `Total Orders` should be calculated as a **distinct count of `Order_ID`** if an order can appear on multiple rows. A simple `COUNTA` is a row count, not necessarily a distinct order count.
- Similarly, `Total Customers` should use a distinct count of `Customer_ID`.
- Profit margin at dashboard level should be calculated as `Total Profit / Total Sales`, rather than taking a simple average of row-level profit margins.
- Review the PivotTable fields and dashboard source ranges after changing the source column names or structure.

## Suggested Repository Structure

```text
Excel-VBA-Automation/
│
├── README.md
├── VBA/
│   ├── Format_Sales_Dataset.bas
│   ├── Sales_Dataset_PivotTables.bas
│   └── Sales_Dashboard_Dataset.bas
│
├── Dataset/
│   └── sales_dataset.xlsx
│
└── Screenshots/
    ├── formatted_dataset.png
    ├── pivot_analysis.png
    └── sales_dashboard.png
```

Adjust the filenames and folders to match the files you actually upload. Avoid committing confidential or personally identifiable business data; use a sample or synthetic dataset when publishing the repository publicly.

## Future Enhancements

- Add slicers for Region, Category, Customer Segment, and Year.
- Connect slicers to compatible PivotTables and PivotCharts.
- Make every KPI card respond to slicer selections.
- Add distinct-count measures for orders and customers where supported.
- Add automated validation and a refresh log.

---

## Project Summary

**Excel VBA Automation** demonstrates how VBA can automate repetitive spreadsheet presentation and reporting tasks: formatting a raw sales dataset, generating multiple PivotTable summaries, and rebuilding a dashboard with KPIs and charts. It is intended as a practical Excel automation and data-analysis portfolio project.
