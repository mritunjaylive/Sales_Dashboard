# 📊 Sales Intelligence Dashboard — Power BI Project

A complete end-to-end Business Intelligence project built using **Power BI Desktop**, covering data modelling, DAX measures, and interactive dashboard development across 5 report pages.

---

## 🗂️ Repository Contents

```
📁 sales-intelligence-dashboard/
│
├── 📄 README.md                    ← You are here
├── 📊 fact_Sales_Raw.xlsx          ← Source data file (raw transactions)
└── 📈 Sales_Data.pbix              ← Power BI Dashboard file
```

---

## 📖 Introduction

This project simulates a real-world sales analytics scenario for a mid-sized technology company selling **Hardware**, **Software**, and **Services** across India and international markets.

### The Problem
The sales team had no single consolidated view of performance. Revenue, targets, product returns, and salesperson performance were scattered across multiple Excel sheets. Monthly reporting took hours of manual work with no guarantee of accuracy.

### The Solution
A fully interactive **5-page Power BI Dashboard** that answers the following questions in seconds:

- How is the business performing against targets?
- Which region and salesperson is performing best and worst?
- Which products are most profitable and most returned?
- Who are the top customers and how concentrated is the revenue?
- How has the business grown year over year?

### What I Learned
- Star Schema data modelling in Power BI
- Power Query for data transformation and dimension table creation
- DAX for writing 16+ business intelligence measures
- Cross-filtering, slicers, conditional formatting and time intelligence
- Translating raw data into a business story with actionable insights

---

## 📦 Data Source

### File: `fact_Sales_Raw.xlsx`

The source file is a single flat Excel table containing **5,528 sales transactions** across 3 years.

| Detail | Value |
|---|---|
| Time Period | January 2022 — December 2024 |
| Total Rows | 5,528 transactions |
| Total Columns | 16 columns |
| Products | 15 (Hardware, Software, Services) |
| Customers | 50 unique customers |
| Salespersons | 15 across 8 regional teams |
| Regions | North India, South India, East India, West India, Americas, Europe, MEA, Asia Pacific |

### Column Descriptions

| Column | Type | Description |
|---|---|---|
| SaleID | Integer | Unique identifier for each transaction |
| SaleDate | Date | Date of the transaction |
| CustomerName | Text | Customer company or individual name |
| CustomerSegment | Text | Enterprise / B2B / SMB / B2C |
| CustomerCity | Text | City where customer is located |
| ProductName | Text | Name of the product sold |
| ProductCategory | Text | Hardware / Software / Services |
| SalespersonName | Text | Name of the salesperson |
| SalespersonTeam | Text | Regional team of the salesperson |
| Region | Text | Geographic region of the customer |
| QuantitySold | Integer | Number of units sold |
| UnitPrice | Decimal | Price per unit at time of sale |
| Discount% | Decimal | Discount applied (0 to 0.25) |
| Revenue | Decimal | Net revenue after discount |
| COGS | Decimal | Cost of Goods Sold |
| ReturnFlag | Boolean | TRUE if order was returned |

### Data Story Elements
The data was designed with embedded story elements to make analysis meaningful:
- **South team** applies higher discounts (10–25%) but still underperforms on revenue
- **Mobile/Hardware** products have elevated return rates (10%+) from 2023 onward
- **Q4 seasonality** — October, November, December are peak months every year
- **YoY growth** — approximately 18% growth from 2022→2023 and 28% from 2023→2024

---

## 🛠️ Steps to Build the Dashboard

### Prerequisites
- Power BI Desktop (free download from [microsoft.com/powerbi](https://www.microsoft.com/en-us/power-platform/products/power-bi/desktop))
- Microsoft Excel (for viewing the source file)

---

### Step 1 - Load the Source File

1. Open Power BI Desktop
2. Click **Home → Get Data → Excel Workbook**
3. Select `fact_Sales_Raw.xlsx`
4. Check the `fact_Sales_Raw` sheet
5. Click **Transform Data** to open Power Query Editor

---

### Step 2 - Fix the Date Column

In Power Query Editor with `fact_Sales_Raw` selected:

1. Click the data type icon on the `SaleDate` column header (shows `123`)
2. Change it to **Date**
3. If prompted — click **Replace Current Step**
4. Dates should now show as `15-Jan-2022` format

---

### Step 3 - Create Dimension Tables

For each dimension table below, follow this pattern:
- Right-click `fact_Sales_Raw` in the left panel → **Duplicate**
- Rename the duplicate
- Select only the relevant columns → right-click → **Remove Other Columns**
- **Home → Remove Rows → Remove Duplicates**
- **Add Column → Index Column → From 1** → rename to the Key column
- Drag the Key column to be first

#### dim_Customer
Keep columns: `CustomerName`, `CustomerSegment`, `CustomerCity`
Add key: `CustomerKey`

#### dim_Product
Keep columns: `ProductName`, `ProductCategory`, `UnitPrice`
Add key: `ProductKey`

#### dim_Salesperson
Keep columns: `SalespersonName`, `SalespersonTeam`
Add key: `SalespersonKey`

#### dim_Region
Keep columns: `Region`
Add key: `RegionKey`

---

### Step 4 - Add Foreign Keys to fact_Sales

For each dimension table, merge it back into the fact table:

1. Select `fact_Sales_Raw` in Power Query
2. **Home → Merge Queries as New**
3. Match on the text column (e.g. `CustomerName` ↔ `CustomerName`)
4. Join Kind: **Left Outer**
5. Expand the merged column — keep only the Key column
6. Rename and combine all 4 merges into a single clean `fact_Sales` table
7. Remove original text columns: `CustomerName`, `CustomerSegment`, `CustomerCity`, `ProductName`, `ProductCategory`, `SalespersonName`, `SalespersonTeam`, `Region`

Click **Home → Close & Apply**

---

### Step 5 - Create dim_Date Using DAX

In Power BI Desktop:

1. Go to **Modeling → New Table**
2. Paste this formula:

```dax
dim_Date = 
ADDCOLUMNS(
    CALENDAR(DATE(2022,1,1), DATE(2024,12,31)),
    "Year",           YEAR([Date]),
    "Month",          MONTH([Date]),
    "MonthName",      FORMAT([Date], "MMMM"),
    "MonthShort",     FORMAT([Date], "MMM"),
    "Quarter",        QUARTER([Date]),
    "QuarterName",    "Q" & QUARTER([Date]),
    "WeekNumber",     WEEKNUM([Date]),
    "DayName",        FORMAT([Date], "DDDD"),
    "IsWeekend",      IF(WEEKDAY([Date],2) >= 6, TRUE(), FALSE()),
    "MonthYear",      FORMAT([Date], "MMM YYYY"),
    "YearMonth",      YEAR([Date])*100 + MONTH([Date])
)
```

3. Go to **Table Tools → Mark as Date Table → select `Date` column**

---

### Step 6 - Create dim_Target Using DAX

**Modeling → New Table**:

```dax
dim_Target = 
DATATABLE(
    "Year", INTEGER, "Month", INTEGER, "TargetRevenue", INTEGER,
    {
        {2022,1,11000000},{2022,2,11500000},{2022,3,17000000},
        {2022,4,12000000},{2022,5,13000000},{2022,6,14500000},
        {2022,7,11500000},{2022,8,12000000},{2022,9,17500000},
        {2022,10,18500000},{2022,11,20000000},{2022,12,14500000},
        {2023,1,12500000},{2023,2,13500000},{2023,3,19000000},
        {2023,4,14000000},{2023,5,15000000},{2023,6,16500000},
        {2023,7,13000000},{2023,8,14000000},{2023,9,20000000},
        {2023,10,21000000},{2023,11,23000000},{2023,12,17000000},
        {2024,1,14500000},{2024,2,15500000},{2024,3,22000000},
        {2024,4,16000000},{2024,5,17500000},{2024,6,19000000},
        {2024,7,15000000},{2024,8,16500000},{2024,9,23000000},
        {2024,10,24500000},{2024,11,26000000},{2024,12,19500000}
    }
)
```

Add a calculated column for the relationship:

```dax
YearMonth = dim_Target[Year] * 100 + dim_Target[Month]
```

---

### Step 7 - Build Relationships in Model View

Click the **Model View** icon (third icon on left sidebar).

Drag and drop to create these relationships:

| From (1 side) | To (* side) | Type |
|---|---|---|
| `dim_Customer[CustomerKey]` | `fact_Sales[CustomerKey]` | One to Many |
| `dim_Product[ProductKey]` | `fact_Sales[ProductKey]` | One to Many |
| `dim_Salesperson[SalespersonKey]` | `fact_Sales[SalespersonKey]` | One to Many |
| `dim_Region[RegionKey]` | `fact_Sales[RegionKey]` | One to Many |
| `dim_Date[Date]` | `fact_Sales[SaleDate]` | One to Many |
| `dim_Date[YearMonth]` | `dim_Target[YearMonth]` | One to Many |

---

### Step 8 - Fix Sort Order for Date Columns

In the Fields panel, for each column below:
1. Click the column
2. **Column Tools → Sort by Column**

| Column | Sort By |
|---|---|
| `dim_Date[MonthName]` | `Month` |
| `dim_Date[MonthShort]` | `Month` |
| `dim_Date[MonthYear]` | `YearMonth` |
| `dim_Date[QuarterName]` | `Quarter` |

---

### Step 9 - Create a Measures Table

**Modeling → New Table**:

```dax
_Measures = ROW("_", BLANK())
```

Hide the `_` column by right-clicking it → **Hide**.

---

### Step 10 — Write DAX Measures

Select the `_Measures` table and use **Home → New Measure** for each:

#### Revenue & Profit

```dax
Total Revenue = SUM(fact_Sales[Revenue])

Total COGS = SUM(fact_Sales[COGS])

Gross Profit = [Total Revenue] - [Total COGS]

Gross Profit Margin % = DIVIDE([Gross Profit], [Total Revenue], 0)

Total Units Sold = SUM(fact_Sales[QuantitySold])

Avg Order Value = DIVIDE([Total Revenue], DISTINCTCOUNT(fact_Sales[SaleID]), 0)

Discount Impact = 
SUMX(fact_Sales,
    fact_Sales[QuantitySold] * fact_Sales[UnitPrice] * fact_Sales[Discount%])

Return Rate % = 
DIVIDE(
    CALCULATE(COUNTROWS(fact_Sales), fact_Sales[ReturnFlag] = TRUE()),
    COUNTROWS(fact_Sales), 0)

Net Revenue = 
CALCULATE([Total Revenue], fact_Sales[ReturnFlag] = FALSE())
```

#### Time Intelligence

```dax
Revenue LY = 
CALCULATE([Total Revenue], SAMEPERIODLASTYEAR(dim_Date[Date]))

Revenue YoY % = 
DIVIDE([Total Revenue] - [Revenue LY], [Revenue LY], BLANK())

Revenue MTD = TOTALMTD([Total Revenue], dim_Date[Date])

Revenue QTD = TOTALQTD([Total Revenue], dim_Date[Date])

Revenue YTD = TOTALYTD([Total Revenue], dim_Date[Date])

Revenue Rolling 3M = 
CALCULATE(
    [Total Revenue],
    DATESINPERIOD(dim_Date[Date], LASTDATE(dim_Date[Date]), -3, MONTH))
```

#### Targets

```dax
Target Revenue = SUM(dim_Target[TargetRevenue])

Revenue vs Target = [Total Revenue] - [Target Revenue]

Target Achievement % = DIVIDE([Total Revenue], [Target Revenue], 0)
```

#### Customer & Salesperson

```dax
Total Customers = DISTINCTCOUNT(fact_Sales[CustomerKey])

Revenue per Customer = DIVIDE([Total Revenue], [Total Customers], 0)

Rep Rank = 
RANKX(
    ALLSELECTED(dim_Salesperson[SalespersonName]),
    [Total Revenue], , DESC, DENSE)

Revenue CAGR = 
VAR yr1 = CALCULATE([Total Revenue], dim_Date[Year] = 2022)
VAR yr3 = CALCULATE([Total Revenue], dim_Date[Year] = 2024)
RETURN IF(yr1 <> 0, POWER(DIVIDE(yr3, yr1, 0), 1/2) - 1, BLANK())
```

---

## 📊 Dashboard Pages

### Page 1 - Executive Summary

**Purpose:** High-level overview of overall business health.

**Visuals:**
- 5 KPI Cards : Total Revenue, Gross Profit, Gross Profit Margin %, Target Achievement %, Return Rate %
- Line Chart : Revenue vs Target by Month/Year
- Bar Chart : Total Revenue by Region
- Table : Top 10 Products with Revenue, Gross Profit Margin % and Return Rate %
- Slicers : Year (dropdown), MonthName (dropdown)

**Key Insight:** Revenue grew consistently across 3 years with clear Q4 seasonality. Target achievement at 76.3% indicates room for improvement.

---

### Page 2 - Sales Performance

**Purpose:** Evaluate individual and team sales performance.

**Visuals:**
- 4 KPI Cards : Total Revenue, Target Revenue, Target Achievement %, Avg Order Value
- Table : Salesperson Leaderboard with data bars and Revenue Rank
- Column Chart : Revenue by Salesperson Team
- Bar Chart : Discount Impact by Salesperson
- Slicers : Year, MonthName

**Key Insight:** South team (Rahul Mehta, Sneha Iyer) has the highest discount impact but the lowest revenue ranking — discounting is not converting to sales performance.

---

### Page 3 - Customer Intelligence

**Purpose:** Understand customer segments, geography and revenue concentration.

**Visuals:**
- 4 KPI Cards : Total Customers, Revenue per Customer, Total Revenue, Return Rate %
- Donut Chart : Revenue by Customer Segment
- Table : Top 10 Customers with Revenue, Gross Profit Margin %, Return Rate %
- Map : Revenue bubbles by Customer City
- Line Chart : Revenue trend by Customer Segment (2022–2024)
- Slicers : Year, CustomerSegment

**Key Insight:** B2B segment drives 51.47% of revenue. Top 10 customers account for 36% of total revenue — a concentration risk that needs active diversification.

---

### Page 4 - Product Analysis

**Purpose:** Analyse product profitability, volume and return rates.

**Visuals:**
- 4 KPI Cards : Total Revenue, Gross Profit Margin %, Return Rate %, Discount Impact
- Stacked Bar / Donut : Revenue by Product Category
- Table : Revenue and Margin by Product with conditional formatting
- Bar Chart : Revenue by Product Name
- Column Chart : Units Sold by Category
- Bar Chart : Return Rate by Category (color coded)
- Slicers : Year, MonthName

**Key Insight:** Hardware drives 84.6% of revenue but Software has 70%+ gross margin. Shifting sales focus toward Software would significantly improve overall profitability.

---

### Page 5 - Trends & Time Analysis

**Purpose:** Understand growth trends, seasonality and year-over-year patterns.

**Visuals:**
- 5 KPI Cards : Total Revenue, Revenue YoY %, Revenue MTD, Revenue YTD, Revenue CAGR
- Line Chart : Year over Year comparison (2022, 2023, 2024 as 3 separate lines)
- Matrix / Heatmap : Revenue by Year × Month (gradient color)
- Bar Chart : Revenue MTD vs Revenue LY by Month
- Column Chart : Quarterly Revenue by Year
- Slicers : MonthName, Year

**Key Insight:** 14.7% CAGR over 3 years confirms consistent growth. October, November and December are peak months every single year — confirmed by both the heatmap and quarterly chart.

---

## 🔍 Key Findings

| # | Finding | Detail |
|---|---|---|
| 1 | Strong Revenue Growth | ₹131M (2022) → ₹197M (2024) — 28% growth, 14.7% CAGR |
| 2 | South Region Problem | Highest discounts, lowest revenue — training and pipeline issue |
| 3 | Software Profit Gap | 70%+ margin but only 10% of revenue — major untapped opportunity |
| 4 | Customer Concentration | B2B = 51% revenue, Top 10 customers = 36% — high dependency risk |
| 5 | Mobile Return Alert | 10%+ return rate on SmartPhone X12 and Tablet ProMax — flag for review |
| 6 | Q4 Seasonality | Oct/Nov/Dec are peak every year — pipeline must start in August |

---

## 🚀 Future Course of Action

1. **Fix South Region** : Implement discount approval process and rep coaching program
2. **Invest in Software** : Realign incentives and training to grow Software revenue share to 25%+
3. **Diversify Customer Base** : No single customer should exceed 10% of total revenue
4. **Investigate Mobile Returns** : Escalate to product team for quality or expectation review
5. **Protect Q4 Pipeline** : Start pipeline building in August every year
6. **Monitor Weekly** : Executive Summary reviewed every Monday, Sales Performance every Friday

---

## 🧰 Tools Used

| Tool | Purpose |
|---|---|
| Microsoft Excel | Source data storage and manual data corrections |
| Power BI Desktop | Data modelling, DAX, dashboard development |
| Power Query (M) | Data cleaning and transformation |
| DAX | Business intelligence measures |

---

## 👤 Author

**Mritunjay Pandey**
- MCA Semester II
- Aryabhatta Knowledge University, Patna

---

## 📝 Notes

- The dataset is synthetically generated with realistic business patterns for educational purposes
- All revenue figures are in Indian Rupees (₹)
- The `.pbix` file requires Power BI Desktop to open (free download from Microsoft)
- Data covers January 2022 to December 2024 (3 full calendar years)
