# Vrinda Store Data Analysis Dashboard

## Project Overview

This project analyzes the annual sales data of Vrinda Store for 2022 using Microsoft Excel.
The dashboard helps understand:

* Customer behavior
* Sales performance
* Order trends
* Top-performing channels and states
* Gender and age analysis

The final output is an interactive Excel dashboard with PivotTables, PivotCharts, and Slicers.

---

## Objective

Vrinda Store wants to create an annual sales report for 2022 so that they can understand their customers and grow more sales in 2023.

---

## Sample Business Questions

1. Compare sales and orders using a single chart.
2. Which month got the highest sales and orders?
3. Who purchased more in 2022 — men or women?
4. What are the different order statuses?
5. Which top states contributed the most sales?
6. Analyze sales by age group and gender.
7. Which category sold the most?
8. Which channel contributed the maximum sales?

---

## Tools Used

* Microsoft Excel
* Data Cleaning Techniques
* Pivot Tables
* Pivot Charts
* Slicers
* Excel Formulas
* Dashboard Design

---

## Dataset Columns

* Order ID
* Customer ID
* Gender
* Age
* Date
* Status
* Channel
* SKU
* Category
* Size
* Quantity
* Currency
* Amount
* Ship City
* Ship State

---

## Steps Used in the Project

### 1. Data Collection
- Imported Vrinda Store sales dataset into Excel.
- Verified all columns and records.

### 2. Data Cleaning
- Enabled filter dropdowns using: `Data → Filter`
- Checked for:
  - Null/blank values
  - Duplicate values
  - Wrong data types
  - Inconsistent text values

### 3. Data Standardization
- Standardized Gender column:
  - `M` → `Men`
  - `W` → `Women`
- Standardized Quantity column:
  - `One` → `1`
  - `Two` → `2`
- Used: `Ctrl + F → Replace`

### 4. Data Preprocessing
- Created new calculated columns:
  - **Age Group Column**: `=IF(E2>=50,"Senior",IF(E2>=30,"Adult","Teenager"))`
  - **Month Column**: `=TEXT(G2,"mmm")`

### 5. Pivot Table Creation
- Created PivotTables for:
  - Orders vs Sales
  - Men vs Women
  - Order Status
  - Top 5 States
  - Age vs Gender
  - Sales Channels

### 6. Pivot Chart Creation
- Created:
  - Combo Chart
  - Pie Chart
  - Bar Chart
  - Column Chart

### 7. Dashboard Design
- Added chart titles
- Removed gridlines
- Formatted charts
- Aligned visuals properly
- Created interactive dashboard layout

### 8. Slicer Integration
- Added slicers for:
  - Month
  - Channel
  - Category
- Connected slicers to all PivotTables using: `Report Connections`

### 9. Dashboard Insights
- Analyzed:
  - Monthly sales trends
  - Gender-wise sales
  - Order status distribution
  - Top-performing states
  - Age group contribution
  - Channel contribution

### 10. Final Dashboard Creation
- Built an interactive Excel dashboard for business decision-making and sales analysis.

---

# Final Dashboard Insights

* Women contributed the highest sales.
* Delivered orders were the majority.
* Ajio/Myntra contributed major sales.
* Certain states generated the highest revenue.
* Adults were the primary buyers.

---

# Dashboard Features

* Interactive dashboard
* Dynamic slicers
* Automatic filtering
* Pivot-based analysis
* Easy refresh with:

  ```text
  Data → Refresh All
  ```

---

# File Saving Recommendation

Save workbook as:

```text
Excel Binary Workbook (*.xlsb)
```

OR

```text
Excel Macro-Enabled Workbook (*.xlsm)
```

This preserves:

* PivotTables
* Slicers
* Charts
* Queries
* Dashboard formatting

---

# Conclusion

This Excel dashboard project demonstrates:

* Data cleaning
* Data preprocessing
* Pivot Table analysis
* Dashboard design
* Business insights generation

The project can be extended further using:

* Power BI
* SQL
* Python
* Advanced Excel automation

---

# Author

Vrinda Store Data Analysis Project
