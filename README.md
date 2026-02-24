<img width="1707" height="685" alt="image" src="https://github.com/user-attachments/assets/a27aa936-65be-412c-be11-10e928d9cba3" />
An interactive Excel dashboard built to answer one question at a glance:
**"Are we performing well this month?"**

![Dashboard Preview](dashboard_preview.png)

---

## 🗂️ Table of Contents

- [Overview](#overview)
- [Dataset](#dataset)
- [Dashboard Features](#dashboard-features)
  - [KPI Cards](#-kpi-cards)
  - [Smart Filter](#-smart-filter)
  - [Trend Chart — Smart Switching](#-trend-chart--smart-switching)
  - [Sales by Region](#-sales-by-region)
  - [Sales by Category](#-sales-by-category)
- [How the Smart Switching Works](#how-the-smart-switching-works)
- [File Structure](#file-structure)
- [How to Use](#how-to-use)
- [Technical Notes](#technical-notes)

---

## Overview

This dashboard was built entirely in **Microsoft Excel** — file is saved in macros, VBA is used in switching monthly and weekly sales trend charts. Apart from that, it uses pure formulas (`SUMIF`, `SUMPRODUCT`, `IF`, `NA()`) combined with carefully structured helper sheets and hand-crafted chart XML to deliver a dynamic, interactive experience.

The dashboard covers **2 years of sales data (2024–2025)** across 5 product categories and 4 regions. All three charts and all three KPI cards update automatically when a filter is applied.

---

## Dataset

The raw data lives in the **`Data`** sheet and contains **1,920 rows** across **7 columns**:

| Column | Description |
|--------|-------------|
| `Month` | Month-year label (e.g. `Jan-24`, `Feb-24` … `Dec-25`) |
| `Week` | Week within the month (`Week 1` – `Week 4`) |
| `Category` | Product category — Electronics, Furniture, Clothing, Food & Beverage, Sports |
| `Region` | Sales region — North, South, East, West |
| `Sales` | Total sales revenue for that row (USD) |
| `Cost` | Cost of goods sold |
| `Profit` | Net profit (`Sales − Cost`) |

> The dataset spans **24 months × 4 weeks × 5 categories × 4 regions = 1,920 rows**.

---

## Dashboard Features

### 💰 KPI Cards

Three cards sit at the top of the dashboard and update automatically based on the active filter:

| KPI | Formula Logic |
|-----|--------------|
| **Total Sales** | Sum of all `Sales` for the selected period |
| **Total Profit** | Sum of all `Profit` for the selected period |
| **Profit Margin %** | `Total Profit ÷ Total Sales` |

- When **"All Year"** is selected → KPIs show full 2024–2025 totals
- When a **specific month** is selected → KPIs show only that month's figures

---

### 🎛️ Smart Filter

Located in the **top-right card** (labelled **SELECT PERIOD**).

- Click the cell inside the filter card to reveal a **dropdown list**
- Options: `All Year` or any of the 24 months (`Jan-24` through `Dec-25`)
- The selected value drives **all three charts and all three KPIs** simultaneously

---

### 📈 Trend Chart — Smart Switching

This is the centerpiece feature of the dashboard. **One chart** that intelligently switches between two views:
<img width="277" height="202" alt="image" src="https://github.com/user-attachments/assets/864d2bee-3c4d-473a-a935-44429c1056cc" /> <img width="728" height="402" alt="image" src="https://github.com/user-attachments/assets/a4f85907-2c9e-4151-bb98-483ce407e9f4" />
#### All Year selected → Monthly Sales Trend
- Plots **24 monthly data points** (Jan-24 through Dec-25)
- Sharp angular line with filled circle markers at each month
- X-axis shows all 24 month labels
- Ideal for spotting year-over-year patterns and seasonal trends

<img width="265" height="203" alt="image" src="https://github.com/user-attachments/assets/6827724c-d199-4268-a2a1-cc3a9998a4de" /> <img width="717" height="392" alt="image" src="https://github.com/user-attachments/assets/0209053b-988c-4fd3-9baf-ffb1ab11036e" />
#### Any month selected → Weekly Sales Trend
- Plots exactly **4 data points** (Week 1, Week 2, Week 3, Week 4)
- Same sharp line + marker style — just zoomed into a single month
- X-axis shows Week 1–4 evenly spaced across the full chart width
- Ideal for understanding within-month performance distribution

The chart title updates dynamically — e.g. `WEEKLY SALES TREND · Jun-24` — so context is always clear.

---

### 🌍 Sales by Region

A **column bar chart** showing sales split across the 4 regions:

- North · South · East · West
- Each region has a distinct color
- Updates to show either full-year totals or a single month depending on the filter

---

### 🏷️ Sales by Category

A **horizontal bar chart** showing sales across the 5 product categories:

- Electronics · Furniture · Clothing · Food & Beverage · Sports
- Each category has a distinct color
- Updates dynamically with the filter — useful for seeing which category dominates in a given month

---

## How the Smart Switching Works

The switching mechanism combines **pure Excel formulas** for the data layer with a **small VBA macro** for the chart visibility toggle. The workbook is saved as `.xlsm` (macro-enabled) to support this.

### The VBA Macro

A lightweight VBA macro handles the **chart switching** — when the filter value changes, it shows the Monthly Trend chart and hides the Weekly Trend chart (or vice versa). This gives an instant, clean transition between the two views without any flicker or overlapping axes.

### The `Chart_Data` sheet

This helper sheet contains two data blocks that feed the two charts:

**Block 1 — Monthly (rows 2–25):**
```
=IF(Dashboard!N11="All Year", SUMIF(...monthly sales...), NA())
```
When "All Year" → returns 24 monthly sales values  
When a month is selected → returns `NA()` for all 24 rows

**Block 2 — Weekly (rows 28–31):**
```
=IF(Dashboard!N11="All Year", NA(), SUMPRODUCT(...weekly sales for selected month...))
```
When "All Year" → returns `NA()` for all 4 rows  
When a month is selected → returns 4 weekly sales values (Week 1–4)

### The `Filter_Data` sheet

A second helper sheet calculates filtered KPIs and chart data for the Category and Region charts using the same `IF("All Year", full-year formula, monthly SUMPRODUCT)` pattern.

---

## File Structure

```
Executive_Sales_Snapshot.xlsx
│
├── Dashboard          ← The visible dashboard (1 sheet)
│   └── N10           ← The filter cell (dropdown: All Year / month)
│
├── Chart_Data         ← Helper sheet for the trend line chart
│   ├── A2:B25        ← Monthly data block (24 rows)
│   └── A28:B31       ← Weekly data block (4 rows)
│
├── Filter_Data        ← Helper sheet for KPIs, category & region charts
│   ├── B2:B4         ← KPI values (Sales, Profit, Margin)
│   ├── A7:B11        ← Category breakdown (5 rows)
│   └── A14:B17       ← Region breakdown (4 rows)
│
└── Data               ← Raw data (1,920 rows × 7 columns)
```

---

## How to Use

1. **Open** `Executive_Sales_Snapshot.xlsm` in Microsoft Excel (2016 or later recommended)
2. When prompted, click **"Enable Content"** or **"Enable Macros"** — this is required for the chart switching to work
3. Navigate to the **Dashboard** sheet — it opens by default
3. To **filter by month**: click the blue cell inside the **SELECT PERIOD** card (top-right), then choose any month from the dropdown
4. To **return to the full year view**: select `All Year` from the dropdown
5. All KPI cards and all three charts update **instantly** — no refresh needed

> **Tip:** The chart title updates dynamically. If you see `WEEKLY SALES TREND · Jun-24`, a month filter is active. If you see `MONTHLY SALES TREND · 2024-2025`, you are in the full-year view.

---

## Technical Notes

- **No external dependencies** — everything runs inside the Excel workbook itself
- **VBA macro** — used specifically for switching between the Monthly and Weekly trend charts based on the filter selection. The rest of the dashboard (KPIs, Category chart, Region chart) updates purely through formulas
- **File format:** `.xlsm` (Excel Macro-Enabled Workbook) — required to store and run the VBA macro
- **Formula engine used:** `SUMIF`, `SUMPRODUCT`, `IF`, `IFERROR`, `NA()`, `SUM`
- **Chart XML** was written directly to ensure correct axis positions, embedded `strCache` labels, and proper chart formatting — guaranteeing labels render correctly on first open
- **Row/column headers** are visible on the Dashboard sheet intentionally to allow easy navigation and cell reference
- The dropdown list for the filter is stored in `Filter_Data!D2:D26` (26 options: "All Year" + 24 months)

> ⚠️ **Important:** Because the file contains a VBA macro, Excel may show a security warning on first open. Click **"Enable Content"** or **"Enable Macros"** to allow the chart switching to work correctly.

---

## Built With

- **Microsoft Excel** — charts, formulas, data validation, VBA macro
- **VBA (Visual Basic for Applications)** — chart switching logic

---

*Dashboard covers Jan 2024 – Dec 2025 · 1,920 data rows · 5 categories · 4 regions*
