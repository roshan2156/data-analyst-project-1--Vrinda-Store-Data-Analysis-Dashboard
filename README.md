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

## 5. Data Analysis

Data analysis was performed using PivotTables, PivotCharts, slicers, and dashboards to understand sales performance and customer behavior.

### Analysis Performed

#### Orders vs Sales Analysis
- Compared total orders and total sales month-wise.
- Created using PivotTable and Combo Chart.

#### Men vs Women Analysis
- Compared sales contribution by men and women customers.
- Created using PivotTable and Pie Chart.

#### Order Status Analysis
- Analyzed:
  - Delivered orders
  - Cancelled orders
  - Returned orders
  - Refunded orders
- Created using PivotTable and Pie Chart.

#### Top 5 States Analysis
- Identified states generating highest sales revenue.
- Created using PivotTable and Bar Chart.

#### Age vs Gender Analysis
- Compared customer age groups with gender contribution.
- Created using PivotTable and Column Chart.

#### Channel Analysis
- Compared sales contribution from different channels such as:
  - Amazon
  - Myntra
  - Flipkart
  - Ajio
  - Meesho
- Created using PivotTable and Pie Chart.

---

### Dashboard Design

The dashboard was designed to display all charts and analysis in one interactive screen.

### What Was Done
- Added chart titles
- Removed gridlines
- Formatted charts
- Adjusted chart alignment
- Created interactive dashboard layout

---

### Slicer Integration

Slicers were added to make the dashboard interactive.

### Slicers Added
- Month
- Channel
- Category

### Feature Used
- Report Connections

---

# Final Dashboard Insights

- Women customers contributed the highest sales.
- Most orders were successfully delivered.
- Amazon, Myntra, and Ajio generated major sales.
- Maharashtra and Karnataka generated high revenue.
- Adult customers were the primary buyers.
- Online channels contributed most of the sales.
- Monthly sales performance remained stable.
- Very few orders were cancelled or refunded.

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

Roshan Chhotulal Patil
