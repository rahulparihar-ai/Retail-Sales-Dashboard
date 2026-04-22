# Section 3 — Power BI Dashboard
## Retail Sales & Forecast BI | Complete Setup Guide

---

## Files in This Folder

| File | Description |
|---|---|
| `DAX_Measures.dax` | All DAX measures — copy-paste into Power BI |
| `PowerBI_Setup_Guide.md` | This step-by-step guide |

---

## STEP 1 — Import Data into Power BI

1. Open **Power BI Desktop** (download free from powerbi.microsoft.com)
2. Click **Get Data → Text/CSV**
3. Browse to: `e:\BEE\Section1_Dataset\retail_sales_dataset.csv`
4. Click **Transform Data** (opens Power Query Editor)

### In Power Query Editor:

| Column | Change Type To |
|---|---|
| Order Date | Date |
| Ship Date | Date |
| Sales | Decimal Number |
| Profit | Decimal Number |
| Discount | Decimal Number |
| Quantity | Whole Number |
| All text columns | Text (already set) |

5. Click **Close & Apply**

---

## STEP 2 — Create the Date Table (DAX)

> [!IMPORTANT]
> A separate Date Table is **mandatory** for time intelligence measures
> like YTD Sales, YoY Growth, MTD Sales, etc. to work correctly.

1. Go to **Modeling tab → New Table**
2. Paste this DAX:

```dax
DateTable =
ADDCOLUMNS(
    CALENDAR(DATE(2021,1,1), DATE(2021,12,31)),
    "Year",          YEAR([Date]),
    "Month Number",  MONTH([Date]),
    "Month Name",    FORMAT([Date], "MMMM"),
    "Month Short",   FORMAT([Date], "MMM"),
    "Quarter",       "Q" & QUARTER([Date]),
    "Quarter Number",QUARTER([Date]),
    "Week Number",   WEEKNUM([Date]),
    "Day Name",      FORMAT([Date], "dddd"),
    "Month-Year",    FORMAT([Date], "MMM YYYY"),
    "Year-Month",    FORMAT([Date], "YYYY-MM")
)
```

3. Press **Enter/Commit**
4. In Fields pane, click **DateTable**
5. Go to **Table tools → Mark as date table → select "Date"**

---

## STEP 3 — Build the Data Model

### Go to Model View (left sidebar — 3rd icon)

You should see two tables: `SalesData` and `DateTable`

### Create Relationship:
1. Drag `DateTable[Date]` **onto** `SalesData[Order Date]`
2. A line (relationship) appears between them
3. Double-click the line to verify:
   - **From**: DateTable → Date (One side `1`)
   - **To**: SalesData → Order Date (Many side `*`)
   - **Cross filter direction**: Single
   - **Active**: Yes ✓

### Final Data Model Diagram:

```
┌─────────────────┐          ┌──────────────────────┐
│   DateTable     │  1    *  │      SalesData        │
│─────────────────│──────────│──────────────────────│
│ Date (PK)       │          │ Order ID              │
│ Year            │          │ Order Date (FK)       │
│ Month Number    │          │ Ship Date             │
│ Month Name      │          │ Customer Name         │
│ Quarter         │          │ Segment               │
│ Month-Year      │          │ Region                │
│ Day Name        │          │ Category              │
└─────────────────┘          │ Sub-Category          │
                             │ Product Name          │
                             │ Sales                 │
                             │ Quantity              │
                             │ Discount              │
                             │ Profit                │
                             └──────────────────────┘
```

---

## STEP 4 — Create Measures Table & DAX Measures

### Create an empty Measures Table:
1. **Modeling → New Table**
2. Paste: `_Measures = {""}`
3. Press Enter

### Add each measure:
1. Right-click `_Measures` in Fields pane → **New Measure**
2. Copy each measure from `DAX_Measures.dax` file
3. Press **Enter** to commit

### All Measures Summary:

| Measure Name | Category | Format |
|---|---|---|
| Total Sales | Basic KPI | Currency |
| Total Profit | Basic KPI | Currency |
| Total Quantity | Basic KPI | Whole Number |
| Total Orders | Basic KPI | Whole Number |
| Profit Margin % | Ratio | Percentage |
| Avg Order Value | Ratio | Currency |
| Discount Rate % | Ratio | Percentage |
| YTD Sales | Time Intelligence | Currency |
| YTD Profit | Time Intelligence | Currency |
| MTD Sales | Time Intelligence | Currency |
| QTD Sales | Time Intelligence | Currency |
| Sales Previous Year | Time Intelligence | Currency |
| YoY Sales Growth % | Time Intelligence | Percentage |
| MoM Sales Growth % | Time Intelligence | Percentage |
| Sales Rank by Category | Ranking | Whole Number |
| Loss Orders Count | Conditional | Whole Number |

---

## STEP 5 — Build Report Pages

### Page 1: Overview Dashboard

| Visual Type | Fields | Position |
|---|---|---|
| Card | Total Sales | Top Left |
| Card | Total Profit | Top Center-Left |
| Card | Profit Margin % | Top Center |
| Card | Total Orders | Top Center-Right |
| Card | YTD Sales | Top Right |
| Clustered Bar Chart | Axis: Category, Values: Total Sales, Total Profit | Center Left |
| Line Chart | X: DateTable[Month-Year], Y: Total Sales, YTD Sales | Center Right |
| Donut Chart | Legend: Segment, Values: Total Sales | Bottom Left |
| Map | Location: State, Size: Total Sales | Bottom Right |

### Page 2: Category Analysis

| Visual Type | Fields |
|---|---|
| Matrix | Rows: Category+Sub-Category, Columns: Region, Values: Sales+Profit |
| Stacked Bar | Axis: Sub-Category, Values: Total Sales, color: Category |
| Scatter Plot | X-axis: Total Sales, Y-axis: Profit Margin %, Details: Sub-Category |
| KPI Card | Value: YoY Sales Growth %, Target: 0 |

### Page 3: Regional View

| Visual Type | Fields |
|---|---|
| Filled Map | Location: State, Color: Total Sales |
| Clustered Column | Axis: Region, Values: Total Sales by Category |
| Table | Region, Total Sales, Total Profit, Profit Margin % |
| Treemap | Group: Region/Segment, Values: Total Sales |

### Page 4: Forecast

| Visual Type | Fields |
|---|---|
| Line Chart with Forecast | X: DateTable[Date] (Month), Y: Total Sales |
| Card | MTD Sales |
| Card | MoM Sales Growth % |

---

## STEP 6 — Enable Forecasting

> [!NOTE]
> Forecasting works **only on Line Charts** with a **continuous date axis**.

### Steps:
1. On Page 4, create a **Line Chart**:
   - X-Axis: `DateTable[Date]` (set to **Month** level)
   - Y-Axis: `Total Sales`

2. Click the chart → go to **Analytics pane** (magnifying glass icon)

3. Expand **Forecast** → click **+ Add**

4. Set these values:

   | Setting | Value |
   |---|---|
   | Forecast length | 3 Months |
   | Confidence interval | 95% |
   | Seasonality | Auto |
   | Ignore last | 0 |

5. Click **Apply** — the forecast line (dashed) with shaded band appears

> [!TIP]
> Change "Forecast length" to **6** for a longer prediction window.
> The shaded band shows the confidence interval (upper/lower bounds).

---

## STEP 7 — Add Slicers & Interactivity

### Add these Slicers on every page:
1. Insert → **Slicer** visual
2. Drag field into Slicer

| Slicer | Field | Style |
|---|---|---|
| Date Range | DateTable[Date] | Between (slider) |
| Category | SalesData[Category] | Dropdown |
| Region | SalesData[Region] | Tile (buttons) |
| Segment | SalesData[Segment] | Tile |

### Sync Slicers across pages:
1. Click a Slicer → **View → Sync Slicers**
2. Check all pages where it should apply

---

## STEP 8 — Format the Report

| Setting | Steps |
|---|---|
| Theme | View → Themes → choose "Executive" or "City Park" |
| Dark mode | View → Themes → Customize → set dark background |
| Page background | Format pane → Canvas background → set color |
| Title text box | Insert → Text box → type dashboard title |
| Logo | Insert → Image → browse college logo |
| Page size | View → Page view → set to 16:9 Widescreen |

---

## STEP 9 — Export for Submission

| Format | Steps |
|---|---|
| PDF | File → Export → Export to PDF |
| PNG (per page) | File → Export → Export to PNG |
| PowerPoint | File → Export → Export to PowerPoint |
| Save .pbix | File → Save As → `Retail_Sales_BI.pbix` |

---

## Slicer + Visual Quick Reference

```
FILTER FLOW:
Date Slicer → [All Visuals on Page]
    ↓
Category Slicer → [Bar Charts, Matrix, Donut]
    ↓
Region Slicer → [Map, Column Chart, Table]
```

---

## Common Errors & Fixes

| Error | Cause | Fix |
|---|---|---|
| YTD Sales returns blank | No Date Table relationship | Mark DateTable as Date Table + create relationship |
| SAMEPERIODLASTYEAR returns blank | Only 2021 data (no 2020) | Use MTD/QTD instead for single-year data |
| Map visual not showing | State names not recognized | Add Country column = "United States" |
| Forecast greyed out | X-axis not continuous date | Use DateTable[Date] not SalesData[Order Date] |

---

*BEE Lab Project — Group 4 | B.Tech First Year*
*Section 3: Power BI Dashboard — Retail Sales & Forecast BI*
