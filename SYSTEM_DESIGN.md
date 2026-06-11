# System Design — Root Sourcing Dashboard for GitHub Pages

## Objective

Create a static interactive dashboard that uses all attached Excel workbooks and can be published directly on GitHub Pages without a backend.

## Architecture

```text
Excel Workbooks
   ↓ tools/convert_excel_to_json.py
Compact JSON dataset
   ↓ fetched by index.html
Interactive Dashboard on GitHub Pages
```

## Data pipeline

### 1. Excel extraction

The converter reads `.xlsx` OpenXML contents directly and exports a compact JSON dataset. It extracts:

- Raw workbook sheet matrices for all sheets.
- Normalized root sourcing transactions from `RegRoot.SK` and `RegRoot.KSN`.
- Monthly KPI rows from `RawKPI.SK` and `RawKPI.KSN`.
- Farmer/broker volume contribution from `VolFarmerContri.SK` and `Volume.KSN`.
- Vendor master rows from `Name.SK` and `Name.KSN`.
- Market series from starch, chip, ethanol, dry pulp, and production sheets.

### 2. Static dataset

The generated dataset is stored at:

```text
data/dashboard_data.json
```

This file is loaded by the browser using `fetch()`.

### 3. Frontend dashboard

`index.html` uses:

- Tailwind CSS for layout and styling.
- Chart.js for interactive charts.
- Native JavaScript for filtering, aggregation, export, pagination, and sheet exploration.

## Dashboard modules

### Overview

- Total Volume (MT)
- Weighted Starch (%)
- Average Price (THB/kg)
- Weighted PPDS
- Unique Vendors
- Monthly volume by plant
- KPI trend: PPDS, price, starch
- Vendor type contribution
- Top provinces and vendors

### Sourcing Details

- Volume by broker/farmer/regular roots
- Area insights by plant, district, and root type
- Paginated transaction table
- CSV export for filtered records

### Market Situation

- Tapioca starch price
- Tapioca chip price
- Ethanol price
- Dry pulp price
- Thailand tapioca production and yield

### Sheet Explorer

- Workbook selector
- Sheet selector
- Search within sheet
- Column display limit
- Full CSV export per selected sheet

### Data Quality

- Required field completeness
- Sheet coverage table
- Notes on calculations and assumptions

## Calculation logic

| Metric | Formula |
|---|---|
| Total Volume | `sum(Volume MT)` |
| Weighted Starch | `sum(Weigh starch) / sum(Volume kg)` |
| Average Price | `sum(Amount) / sum(Volume kg)` |
| Weighted PPDS | `sum(Amount) / sum(Weigh starch)` |
| Unique Vendors | Distinct vendor names after filters |

## GitHub Pages deployment

The dashboard is designed as a single-page static site. Required files for deployment:

```text
index.html
data/dashboard_data.json
tools/convert_excel_to_json.py
README.md
SYSTEM_DESIGN.md
```

No server, database, API key, or build tool is required.
