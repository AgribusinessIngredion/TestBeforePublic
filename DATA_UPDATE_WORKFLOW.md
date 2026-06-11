# Data Update Workflow

## Recommended monthly workflow

```text
Excel source files
  ↓
tools/convert_excel_to_json.py
  ↓
data/dashboard_data.json
  ↓
data/csv/*.csv and data/csv/by_sheet/*.csv
  ↓
tools/validate_data.py
  ↓
Commit / Push to GitHub
  ↓
GitHub Pages dashboard update
```

## Step-by-step

### 1. Update source Excel files

Place the latest source files in:

```text
data/source_excel/
```

Keep these file names unchanged unless you also update `SOURCE_FILES` in `tools/convert_excel_to_json.py`:

```text
KSN Root Sourcing report App.xlsx
SK Root Sourcing report App.xlsx
Market situation.xlsx
ST_PPDS_Price template.xlsx
```

### 2. Convert data

```bash
python3 tools/convert_excel_to_json.py
```

The converter creates:

```text
data/dashboard_data.json
data/manifest.json
data/quality_report.json
data/csv/*.csv
data/csv/by_sheet/*/*.csv
```

### 3. Validate PPDS and missing values

```bash
python3 tools/validate_data.py
```

Review these files:

```text
data/quality_report.json
data/ppds_missing_records.csv
data/critical_missing_ppds_records.csv
```

### 4. Publish to GitHub

```bash
git add .
git commit -m "Update dashboard data"
git push
```

## PPDS rule

Recommended interpretation:

```text
PPDS missing + Volume MT = 0 + Amount = 0
→ Keep as audit record, not critical

PPDS missing + Volume MT > 0 or Amount > 0
→ Critical; review source Excel before publishing
```

## Excel template rules

To prevent data loss, keep these rules:

1. Do not rename sheets.
2. Do not rename columns.
3. Do not merge cells inside raw tables.
4. Do not add multi-row headers inside raw tables.
5. Use stable Month values: Jan, Feb, Mar, Apr, May, Jun, Jul, Aug, Sep, Oct, Nov, Dec.
6. Use numeric values only for Volume, Price, Amount, PPDS, and Starch.
7. Do not type units inside numeric cells.
8. Use blank cells for true missing data; do not use `-` or `N/A`.
