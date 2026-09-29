# 🍔 Swiggy Food Delivery Sales Analysis (Excel Project)

An end-to-end Excel analytics project on **197,430 Swiggy food orders** across **28 states/cities in India** (Jan – Aug 2025). The workbook covers data preparation with formulas, analysis with PivotTables, and an interactive dashboard built with charts, map charts and slicers.

---

## 📌 Table of Contents

1. [Project Overview](#-project-overview)
2. [Business Objectives](#-business-objectives)
3. [Dataset Description](#-dataset-description)
4. [Workbook Structure](#-workbook-structure)
5. [Data Preparation](#-data-preparation)
6. [KPIs & Metrics](#-kpis--metrics)
7. [Analysis & Key Insights](#-analysis--key-insights)
8. [Dashboard](#-dashboard)
9. [Excel Skills & Features Used](#-excel-skills--features-used)
10. [How to Use](#-how-to-use)
11. [Limitations & Future Improvements](#-limitations--future-improvements)
12. [Repository Structure](#-repository-structure)

---

## 📖 Project Overview

Food delivery platforms generate huge volumes of order data. This project turns raw Swiggy order records into business insights on **sales, customer ratings, geography, time trends and food preferences**, all inside Microsoft Excel.

| Item | Detail |
|---|---|
| Tool | Microsoft Excel (PivotTables, Charts, Slicers, Map Charts, Formulas) |
| Records | 197,430 orders |
| Period | 1 Jan 2025 – 31 Aug 2025 |
| Coverage | 28 states, 28 cities, 993 restaurants |
| Total Sales | ₹5.30 Crore (₹53,012,506) |

---

## 🎯 Business Objectives

- Measure overall performance (sales, orders, average order value, ratings).
- Identify **top-performing states and cities**.
- Understand **monthly, weekly, daily and quarterly** sales patterns.
- Compare **Veg vs Non-Veg** demand.
- Enable drill-down by **month, restaurant and food category** through an interactive dashboard.

---

## 🗂 Dataset Description

**Sheet:** `Swiggy Data` (Excel Table: `Table1`)

| Column | Type | Description |
|---|---|---|
| State | Text | State of the restaurant |
| City | Text | City of the restaurant |
| Order Date | Date | Date of the order |
| Day | Formula | Day of the week (Mon–Sun) |
| Qruarter | Formula | Quarter (Q1–Q3) derived from the order date |
| Week | Formula | Week number of the year |
| Restaurant Name | Text | Restaurant name |
| Location | Text | Locality/area within the city |
| Category | Text | Menu category (e.g. Recommended, Main Course, Desserts) |
| Dish Name | Text | Name of the dish ordered |
| Food Type | Formula | Veg / Non-veg classification |
| Price (INR) | Number | Price of the dish in ₹ |
| Rating | Number | Average dish rating |
| Rating Count | Number | Number of ratings received |

**Quick data profile**

- 197,430 rows, no blank cells in any column
- 28 states/cities · 993 restaurants · 4,972 unique menu categories
- Price range: ₹0.95 – ₹8,000
- Date range: 01-Jan-2025 to 31-Aug-2025

> Note: the quarter column header is spelled `Qruarter` in the workbook. It is kept as-is here so it matches the PivotTables.

---

## 📁 Workbook Structure

| Sheet | Purpose |
|---|---|
| **Swiggy Data** | Cleaned raw data with derived columns (Day, Quarter, Week, Food Type) |
| **Analysis** | KPI block and all PivotTables that feed the charts |
| **Dashboard** | Final interactive dashboard with charts, map charts and slicers |
| **Sheet3** | Helper pivot (monthly sales) |

---

## 🧹 Data Preparation

Derived columns were created using Excel formulas on the table:

```excel
Day        =TEXT([@[Order Date]],"ddd")
Quarter    ="Q"&INT((MONTH([@[Order Date]])-1)/3)+1
Week       =WEEKNUM([@[Order Date]])
```

**Food Type classification** (keyword-based on the dish name):

```excel
=IF(OR(ISNUMBER(SEARCH("chicken",J2)),
       ISNUMBER(SEARCH("egg",J2)),
       ISNUMBER(SEARCH("fish",J2)),
       ISNUMBER(SEARCH("mutton",J2)),
       ISNUMBER(SEARCH("prawns",J2)),
       ISNUMBER(SEARCH("biryani",J2)),
       ISNUMBER(SEARCH("kebab",J2))),"Non-veg","Veg")
```

Data is stored as an **Excel Table**, so formulas and PivotTables update automatically when rows are added.

---

## 📊 KPIs & Metrics

| KPI | Value |
|---|---|
| **Total Sales** | ₹53,012,506 (≈ ₹5.30 Cr) |
| **Total Orders** | 197,430 |
| **Average Order Value (AOV)** | ₹268.51 |
| **Average Rating** | 4.34 |
| **Total Rating Count** | 5,591,574 |

`AOV = Total Sales ÷ Number of Orders` (calculated with `GETPIVOTDATA`).

---

## 🔍 Analysis & Key Insights

### 1. Monthly Trend (Jan – Aug 2025)
- Sales are fairly stable at roughly ₹6.3M – ₹6.8M per month.
- **Highest:** January (₹6.83M) · **Lowest:** February (₹6.27M, the shortest month).

### 2. Quarterly Analysis
| Quarter | Sales (₹) | Orders | Avg Rating |
|---|---|---|---|
| Q1 | 19,667,822 | 73,096 | 4.34 |
| Q2 | 19,902,257 | 74,163 | 4.34 |
| Q3* | 13,442,427 | 50,171 | 4.34 |

\*Q3 contains only July and August, so it is a partial quarter and not directly comparable. Ratings are consistent across quarters.

### 3. Day-of-Week Trend
- **Saturday** has the highest sales (₹7.78M), followed by Thursday and Sunday.
- **Tuesday** is the weakest day (₹7.36M).
- The gap between best and worst day is small (~5.7%), so demand is evenly spread through the week.

### 4. Veg vs Non-Veg
| Type | Sales | Share |
|---|---|---|
| Veg | ₹35.25M | ~66.5% |
| Non-veg | ₹17.77M | ~33.5% |

Veg dishes account for about two-thirds of sales value.

### 5. Geographic Analysis
- **Top 5 cities by sales:** Bengaluru (₹5.46M), Lucknow (₹3.12M), Hyderabad (₹3.02M), Mumbai (₹3.02M), New Delhi (₹2.83M).
- These five cities contribute **~33%** of total sales (₹17.44M).
- **Bengaluru alone contributes ~10.3%**, almost double the next city.
- Lowest-sales states: Nagaland and Sikkim (< ₹0.6M), indicating room for growth in the North-East.

### 6. Weekly Trend
- Weekly sales stay consistently around ₹1.5M per week.
- Week 1 and week 36 are lower because they are partial weeks at the start and end of the dataset.

---

## 📈 Dashboard

The **Dashboard** sheet includes:

- **KPI cards**: Total Sales, Orders, AOV, Average Rating
- **Line charts**: monthly and weekly sales trend
- **Doughnut charts**: Veg vs Non-Veg split
- **Bar charts**: day-of-week, quarterly and Top 5 cities
- **Map charts (Region Map)**: state-wise sales visualization
- **Slicers**: filter everything by **Month**, **Restaurant Name** and **Category**

> 💡 Add a screenshot at `images/dashboard.png` and it will appear below:

![Dashboard Preview](images/dashboard.png)

---

## 🛠 Excel Skills & Features Used

- PivotTables & PivotCharts
- Slicers (connected to multiple PivotTables)
- Excel Tables (structured references)
- Formulas: `IF`, `OR`, `SEARCH`, `ISNUMBER`, `TEXT`, `WEEKNUM`, `MONTH`, `INT`, `GETPIVOTDATA`
- Charts: line, bar, doughnut, region map
- KPI cards and dashboard layout design
- Grouping dates by month/quarter

---

## ▶️ How to Use

1. Clone or download this repository.
2. Open `Swiggy_Data_Excel.xlsx` in **Microsoft Excel 2016 or later** (Excel 365 recommended, since map charts and slicers need it).
3. Go to the **Dashboard** sheet and use the slicers to filter by month, restaurant or category.
4. Go to the **Analysis** sheet to see the underlying PivotTables.
5. If you edit the data, use **Data → Refresh All** to update the pivots.

---

## ⚠️ Limitations & Future Improvements

**Current limitations**
- Food Type is inferred from keywords, so some dishes may be misclassified (e.g. a "Veg Egg-less Biryani").
- 79,094 rows (~40%) have a Rating Count of 0, meaning the rating is not backed by reviews.
- Each row is a dish listing, so "orders" here means dish records rather than unique customer orders.
- Data covers only Jan–Aug 2025.
- Spelling in headers (`Qruarter`, `Sate`) can be corrected for a cleaner look.

**Future improvements**
- Add Top 10 restaurants and Top 10 dishes analysis
- Rating vs Price correlation analysis
- Power Query for automated cleaning
- Rebuild the dashboard in Power BI or Tableau
- Add a category-level profitability/price band analysis

---

## 📂 Repository Structure

```
swiggy-excel-analysis/
│
├── Swiggy_Data_Excel.xlsx     # Main Excel project file
├── README.md                  # Project documentation
└── images/
    └── dashboard.png          # Dashboard screenshot
```

> ⚠️ The Excel file is ~23 MB. GitHub allows files up to 100 MB, but if you hit issues consider [Git LFS](https://git-lfs.com/).

---

## 👤 Author

PADIGIREDDY BHANUPRKASH REDDY

📧 bhanuprakashreddy.p1019@gmail.com        🔗 [LinkedIn] https://www.linkedin.com/in/padigireddy-bhanuprakash-reddy-476a32275/?isSelfProfile=true 

💻 [GitHub] https://github.com/bhanuprakash1019

---

⭐ If you found this project useful, please give it a star!
