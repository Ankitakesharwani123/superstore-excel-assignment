# Superstore Sales — Excel Practical Assignment

**Author:** Ankita Kesharwani
**Course context:** Excel Practical Assignment (Module 2)
**Dataset:** Superstore Sales (9,994 orders, 2014–2017 — Orders, Returns, People sheets)

This repo contains a fully worked Excel Practical Assignment covering all 11 sections of the
brief, built entirely on live formulas (SUM, SUMIFS, VLOOKUP, DATEDIF, IFERROR, etc.) against
the real Superstore dataset — nothing is hardcoded, so every number recalculates if the
source data changes.

## 📁 Repository structure

```
├── raw-data/
│   └── Superstore_Sales_raw.xlsx          # Original, untouched dataset
├── working-data/
│   └── Superstore_Excel_Assignment.xlsx   # Completed assignment (11 sections + Dashboard)
├── screenshots/
│   └── dashboard_screenshot.png           # Dashboard sheet screenshot
├── docs/
│   └── Superstore_Assignment_Summary.docx # 1-page project summary
└── README.md
```

## ✅ Sections completed (`working-data/Superstore_Excel_Assignment.xlsx`)

| Sheet | Covers |
|---|---|
| `S1_Basics` | SUM, MAX, MIN, SMALL, LARGE, MEDIAN · manual vs `TRANSPOSE()` · COUNT/COUNTA/COUNTBLANK · Status Bar · YoY Growth % & `POWER()` (CAGR) |
| `S2_Formatting` | Currency & conditional cell colour · custom number formats (%, thousands, red negatives, decimals, scientific) · Wrap Text / Merge & Center · Format Painter |
| `S3_SortFilter` | Sort by date, by amount, alphabetically (A-Z/Z-A), by cell colour, 2-level sort |
| `S4_Dates` | Duration between dates/timestamps · `DATEDIF` · `TODAY`, `DATE`, `YEAR`, `MONTH`, `DAY` |
| `S5_Logical` | `IF`, `AND`, `OR`, combined logical tests, `IFERROR` |
| `S6_DataCleaning` | Remove Duplicates · `UPPER/PROPER/LOWER/TRIM/VALUE/LEN/LEFT/RIGHT/MID` · `SEARCH` vs `FIND` · Go To Special · Text to Columns · `CONCATENATE` vs `&` · Find & Replace |
| `S7_PivotTable` | Category / monthly / regional pivot-style summaries · blank-cell grouping · custom groups + conditional formatting |
| `S8_Lookup` | `VLOOKUP` exact & approximate match · named range lookup · `VLOOKUP` vs `HLOOKUP` |
| `S9_ConditionalAgg` | `COUNTIF`, `SUMIF`, `AVERAGEIF` · `COUNTIFS`, `SUMIFS` (multi-criteria) |
| `S10_CondFormat` | Rule-based cell colour · colour scales · data bars · icon sets |
| `S11_Consolidation` | Data ▸ Consolidate (SUM) across sheets · `SUBTOTAL` on filtered data |
| `Dashboard` | Live KPI cards (Total Sales, Profit, Orders, Avg Discount) + bar / pie / line charts |

Every "native Excel step" (e.g. how to actually insert a PivotTable, use Format Painter, or
run Text to Columns in the Excel UI) is documented as a note directly next to the relevant
formulas on each sheet, so the workbook doubles as a walkthrough.

## 📊 Dashboard preview

See `screenshots/dashboard_screenshot.png` — Total Sales $2,297,201 · Total Profit $286,397 ·
9,994 orders · 15.6% avg. discount, with Sales-by-Category, Sales-by-Region, and Monthly
Trend (2017) charts.

## 🧮 How it was verified

Every formula was recalculated with LibreOffice headless and checked for zero formula errors
before being committed — the workbook opens clean in Excel with no `#REF!`, `#N/A`, or
`#DIV/0!` cells.

# Excel Assignment 1 – Section 1: Basics of Excel

## 📊 Project Overview

This project demonstrates the use of basic Excel functions to analyze **monthly sales data**. The objective is to practice fundamental Excel arithmetic and statistical functions such as `SUM`, `MAX`, `MIN`, `SMALL`, `LARGE`, and `MEDIAN`.

The project is based on the **Superstore Sales dataset** and focuses on calculating and interpreting monthly sales figures.

## 🎯 Objectives

* Calculate total annual sales using the `SUM` function.
* Identify the month with the highest sales using `MAX`.
* Identify the month with the lowest sales using `MIN`.
* Find the 2nd lowest sales value using `SMALL`.
* Find the 3rd highest sales value using `LARGE`.
* Calculate the median monthly sales using `MEDIAN`.

## 📁 Dataset

**Dataset:** Superstore Sales

The analysis uses monthly sales figures derived from the Superstore Sales dataset.

### Main Fields Used

* Order Date
* Sales

The monthly sales data is grouped by month and used for the calculations in this assignment.

## 🧮 Excel Functions Used

| Task              | Excel Function | Purpose                               |
| ----------------- | -------------- | ------------------------------------- |
| Total Sales       | `SUM()`        | Calculates total monthly sales        |
| Highest Sales     | `MAX()`        | Finds the highest sales value         |
| Lowest Sales      | `MIN()`        | Finds the lowest sales value          |
| 2nd Lowest Sales  | `SMALL()`      | Finds the second-smallest sales value |
| 3rd Highest Sales | `LARGE()`      | Finds the third-largest sales value   |
| Median Sales      | `MEDIAN()`     | Calculates the middle sales value     |

## 📌 Tasks Completed

### Task 1 – Total Sales

The `SUM` function is used to calculate the total sales for the year.

```excel
=SUM(B2:B13)
```

### Task 2 – Highest Sales Month

The `MAX` function is used to identify the highest monthly sales value.

```excel
=MAX(B2:B13)
```

### Task 3 – Lowest Sales Month

The `MIN` function is used to identify the lowest monthly sales value.

```excel
=MIN(B2:B13)
```

### Task 4 – 2nd Lowest Sales

The `SMALL` function is used to find the second-lowest monthly sales value.

```excel
=SMALL(B2:B13,2)
```

### Task 5 – 3rd Highest Sales & Median

The `LARGE` function is used to find the third-highest monthly sales value.

```excel
=LARGE(B2:B13,3)
```

The `MEDIAN` function is used to calculate the median monthly sales.

```excel
=MEDIAN(B2:B13)
```

## 📊 Project Contents

The Excel workbook contains:

* Monthly sales dataset
* Sales calculations
* Excel formulas
* Highest and lowest sales analysis
* 2nd lowest sales analysis
* 3rd highest sales analysis
* Median sales calculation
* Monthly sales visualization

## 🛠️ Tools & Skills

* Microsoft Excel
* Data Analysis
* Basic Excel Functions
* `SUM`
* `MAX`
* `MIN`
* `SMALL`
* `LARGE`
* `MEDIAN`
* Data Visualization

## 📈 Learning Outcome

Through this project, I practiced using fundamental Excel functions to analyze sales data and extract meaningful information from a dataset.

This project is part of my **Excel Data Analysis learning and portfolio projects**.

## 👩‍💻 Author

**Ankita Kesharwani**

Data Science & AI | Data Analytics | Excel | SQL | Python | Power BI

