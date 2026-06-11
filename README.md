# Root Sourcing & Market Situation Dashboard

Interactive static dashboard for GitHub Pages. The dashboard reads `data/dashboard_data.json` and includes CSV backups for auditing each source sheet.

## Folder structure

```text
root_sourcing_dashboard/
├── index.html
├── README.md
├── SYSTEM_DESIGN.md
├── DATA_UPDATE_WORKFLOW.md
├── update_data.sh
├── update_data.bat
│
├── data/
│   ├── dashboard_data.json                  # Main file used by dashboard
│   ├── manifest.json                        # Data version, file list, row counts
│   ├── quality_report.json                  # Data quality checks
│   ├── ppds_missing_records.csv             # Rows with missing PPDS
│   ├── ppds_zero_records.csv                # Rows with PPDS = 0
│   ├── critical_missing_ppds_records.csv     # Critical rows requiring review
│   │
│   ├── source_excel/                        # Original Excel files
│   │   ├── KSN Root Sourcing report App.xlsx
│   │   ├── SK Root Sourcing report App.xlsx
│   │   ├── Market situation.xlsx
│   │   └── ST_PPDS_Price template.xlsx
│   │
│   └── csv/
│       ├── sourcing_records_all.csv
│       ├── sk_regroot.csv
│       ├── ksn_regroot.csv
│       ├── kpi_monthly_all.csv
│       ├── sk_kpi_monthly.csv
│       ├── ksn_kpi_monthly.csv
│       ├── volume_contribution_all.csv
│       ├── vendor_master_all.csv
│       ├── sheet_index.csv
│       └── by_sheet/                        # CSV export for every Excel sheet
│
├── tools/
│   ├── convert_excel_to_json.py             # Standard-library converter
│   └── validate_data.py                     # Data quality validator
│
└── .github/workflows/
    └── build-data.yml                       # Optional GitHub Actions automation
```

## Run locally

```bash
python3 tools/convert_excel_to_json.py
python3 tools/validate_data.py
python3 -m http.server 8000
```

Open:

```text
http://localhost:8000
```

On Windows:

```bat
update_data.bat
```

## Update data in the future

1. Replace or update the Excel files in `data/source_excel/`.
2. Run `python3 tools/convert_excel_to_json.py`.
3. Check `data/quality_report.json`.
4. Commit and push to GitHub.
5. GitHub Pages will show the updated dashboard.

## Current data quality

- Root sourcing records: 14,875 rows
- KPI rows: 96 rows
- Volume rows: 96 rows
- Vendor master rows: 1,694 rows
- Workbook sheets exported: 75 sheets
- PPDS complete: 14,768 rows
- PPDS missing: 107 rows
- Critical missing PPDS: 0 rows

The current missing PPDS rows are zero-volume / zero-amount audit records, so they are kept in the dataset but excluded from average PPDS calculations in the dashboard.
