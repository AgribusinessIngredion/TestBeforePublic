# System Design

## Goal

Create a GitHub Pages-ready static dashboard that can display sourcing, KPI, volume, market, and raw sheet data without requiring a backend server.

## Architecture

```text
Source Excel files
  └── data/source_excel/*.xlsx
        ↓
Python converter
  └── tools/convert_excel_to_json.py
        ↓
Main data file
  └── data/dashboard_data.json
        ↓
Interactive dashboard
  └── index.html + Chart.js + Tailwind CSS
```

## Data layers

### 1. Source layer

Original workbooks are stored in:

```text
data/source_excel/
```

### 2. Audit layer

Every relevant table and every original sheet is exported to CSV:

```text
data/csv/
data/csv/by_sheet/
```

This prevents data loss because any missing value can be traced back to a source workbook and sheet.

### 3. Application layer

The dashboard reads:

```text
data/dashboard_data.json
```

This single JSON file is faster and safer for GitHub Pages than loading multiple Excel files in the browser.

## Dashboard modules

- Executive KPI cards
- Plant comparison: Sikhiu vs Kalasin
- PPDS / Price / Starch trend
- Volume contribution by vendor type
- Market situation charts
- Sheet explorer for all workbooks and all sheets
- Data quality panel
- Searchable detail table

## Data quality design

The system separates PPDS data into four classes:

```text
Complete PPDS
Missing PPDS
Zero PPDS
Critical missing PPDS
```

Critical missing PPDS means PPDS is blank while volume, amount, or price exists. These rows should be reviewed before publishing.

## Why JSON + CSV instead of Excel in browser

The dashboard does not read Excel directly because browser-side Excel loading is slower and more fragile. JSON is used for dashboard performance, while CSV is used for audit and manual review.
