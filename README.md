# 📊 E-Commerce Sales Analytics — Excel Capstone Project

> **Data Analytics Module-End Assignment | Data Cleaning, Analysis & Visualization with Excel**

---

## 🗂️ Project Overview

This project involves end-to-end data analysis of an E-Commerce Sales dataset using Microsoft Excel. The work covers the full analytics pipeline — from raw data import and cleaning, through formula-based analysis and statistical summarization, to interactive dashboards and business insights.

**Dataset:** E-Commerce Sales Dataset (~2,000 sales transactions across 500 customers, 100 products, and 20 stores)

---

## 📸 Dashboard Preview

> Built entirely in Microsoft Excel using Pivot Charts, Slicers, and PivotTables.

![E-Commerce Sales Analytics Dashboard](screenshots/dashboard.png)

**What's visible in the dashboard:**
- **Revenue by Category** — Clustered bar chart comparing Total Revenue, Profit, Average Order Value, and Count across 5 categories
- **Payment Type** — Pie chart showing PayPal (35%), COD (33%), Credit Card (32%) split
- **Revenue by Store Type** — Bar chart: Online ($914K) leads, followed by Flagship ($848K) and Outlet ($411K)
- **Revenue by Gender** — Horizontal bar comparing Male vs Female revenue and profit
- **Sum by Loyalty Level** — Platinum customers dominate at $656K, followed by Bronze ($522K), Gold ($517K), Silver ($460K)
- **Revenue by Region & Gender** — 3D clustered chart breaking down East, North, South, West by gender

---

## 🏗️ Dataset Structure (Star Schema)

The project uses a **dimensional data model** with 4 tables:

| Table | Description | Records |
|---|---|---|
| `Customer_Dim` | Customer demographics, location, loyalty level | 500 customers |
| `Product_Dim` | Product info, category, brand, cost, stock | 100 products |
| `Store_Dim` | Store name, region, city, store type | 20 stores |
| `Sales_Fact` | Orders with quantity, price, discount, payment | 2,000 transactions |

---

## 📋 Excel Workbook Sheet Index

| Sheet | Purpose |
|---|---|
| `Customer_Dim` | Raw + cleaned customer data with TRIM, PROPER, IF formulas |
| `Product_Dim` | Raw + imputed product cost and stock using AVERAGEIF |
| `Store_Dim` | Store reference table |
| `Sales_Fact` | Raw + imputed sales fields using IF-based formulas |
| `ALL DATASETS WITH V&XLOOKUP` | Master table joined using VLOOKUP & XLOOKUP |
| `Final Clean Datasets` | Final cleaned dataset with all features merged |
| `Formula sheet` | SUM, AVERAGE, COUNTA, SUMIF, COUNTIF calculations |
| `QUIKAnalysis` | Pivot table — category-wise revenue breakdown |
| `Pivot Tables` | Multi-dimensional pivot summaries |
| `DASHBOARD` | Interactive E-Commerce Sales Analytics Dashboard |
| `CategoryProfitAnalysis` | Profit breakdown by product category |
| `Overall_Analysis` | Aggregated analysis by Category, Payment, Region, Loyalty |

---

## 🔧 Skills Demonstrated

### 1. 📥 Data Entry & Organization
- Imported and structured 4 separate dimension and fact tables
- Applied proper data types and formatted tables in Excel
- Organized workbook using multiple named sheets/tabs

### 2. 🧹 Data Cleaning & Transformation
- **Duplicates:** Used MATCH function to detect and flag duplicate Customer IDs
- **Missing values:** Imputed missing `Cost` and `Stock` using `AVERAGEIF` by sub-category; imputed missing `Quantity`, `Unit Price`, and `Discount` using reverse calculation formulas
- **Name formatting:** Applied `TRIM(PROPER())` to standardize customer names
- **Null handling:** Used `IF(ISBLANK(), ...)` to replace blanks with "Unknown" for Loyalty Level
- **Irrelevant data:** Identified and removed non-contributing columns

### 3. 🔗 Table Joining (VLOOKUP & XLOOKUP)
- Used **VLOOKUP** to pull `Cost` and `Sub_Category` from `Product_Dim`
- Used **XLOOKUP** to fetch `Region`, `Store_Type`, `City`, `Gender`, `Loyalty_Level`, `State`, `Category`, `Product_Name`, `Stock` across dimension tables
- Computed **Profit** = `Total Amount − (Cost × Quantity)` in the joined master table

### 4. 📐 Excel Formulas & Functions

| Formula | Usage |
|---|---|
| `SUM` | Total Revenue, Total Profit, Total Quantity |
| `AVERAGE` | Average Order Value |
| `COUNTA` | Total number of orders |
| `SUMIF` | Revenue by category (Sports, Electronics, Home, Clothing, Beauty) |
| `COUNTIF` | Order count by category |
| `IF` | Classify orders as "Bulk order" (qty ≥ 3) vs "Single order" |
| `Nested IF / IFS` | Age group segmentation |
| `AVERAGEIF` | Category-wise average cost/stock imputation |
| `TRIM / PROPER` | Text standardization |
| `ISBLANK` | Null value detection |
| `INDEX / MATCH` | Lookup demonstrations |

### 5. 📊 Descriptive Statistics
- Computed Mean, Median, Mode, Standard Deviation using Excel formulas
- Used the **Analysis ToolPak** for comprehensive descriptive statistics on sales, profit, quantity, discount fields

### 6. 🔄 Pivot Tables
- Created pivot tables for category-wise revenue summarization
- Analyzed `Sum of Unit Price`, `Sum of Discount`, `Sum of Total Amount` by category
- Multi-dimensional breakdowns across Region, Payment Type, and Loyalty Level

### 7. 📈 Data Visualization & Dashboard
- Built an **Interactive E-Commerce Sales Analytics Dashboard** (`DASHBOARD` sheet)
- Chart types used: Bar charts, Column charts, Pie charts, PivotCharts
- Visualized: Category revenue, regional performance, payment type distribution, loyalty segment analysis

---

## 📊 Key Findings

### 💰 Revenue by Category (Total: $2,174,394)

| Category | Total Sales | Total Profit |
|---|---|---|
| 🥇 Sports | $542,010 | — |
| 🥈 Electronics | $504,311 | $127,560 |
| 🥉 Home | $427,296 | $192,592 |
| Clothing | $379,216 | $146,299 |
| Beauty | $321,561 | $134,708 |

### 💳 Revenue by Payment Type

| Payment Type | Total Sales | Total Profit |
|---|---|---|
| PayPal | $756,755 | $318,804 |
| COD | $702,459 | $274,364 |

### 🗺️ Revenue by Region

| Region | Total Sales | Total Profit |
|---|---|---|
| South | $635,213 | $272,403 |
| East | $622,404 | $253,505 |
| West | $270,676 | $111,800 |

### 👑 Revenue by Loyalty Level

| Loyalty Level | Total Sales | Total Profit |
|---|---|---|
| Platinum | $656,104 | $277,582 |
| Silver | $460,533 | $185,858 |
| Bronze | $522,623 | $203,625 |

---

## 🛠️ Tools Used

- **Microsoft Excel** — Primary tool for all analysis
- **Pivot Tables** — Data summarization
- **Analysis ToolPak** — Descriptive statistics
- **Excel Charts & PivotCharts** — Data visualization
- **VLOOKUP / XLOOKUP** — Table joining
- **IF / IFS / AVERAGEIF / SUMIF / COUNTIF** — Conditional logic and aggregation

---

## 📁 Project Structure

```
📦 Excel-Ecommerce-Capstone/
├── 📊 ARUN_CAPSTONE_PROJECT.xlsx     ← Main Excel workbook (all sheets)
└── 📄 README.md                      ← Project documentation
```

---

## 🎯 Assignment Requirements Fulfilled

| Requirement | Status |
|---|---|
| Data entry, import, and table organization | ✅ Done |
| Remove duplicates & handle missing/inconsistent data | ✅ Done |
| Remove irrelevant columns | ✅ Done |
| Excel formulas (SUM, IF, COUNTIF, TEXT, AVERAGE) | ✅ Done |
| VLOOKUP / XLOOKUP to combine tables | ✅ Done |
| Pivot Tables for summarization | ✅ Done |
| Descriptive statistics (mean, median, SD, Analysis ToolPak) | ✅ Done |
| ≥3 chart types + interactive dashboard | ✅ Done |
| One-page summary of key findings | ✅ Done |

---

## 👤 Author

**ARUN C**
Data Analytics Learner | Excel Capstone Project
