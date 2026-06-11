# Root Sourcing & Market Situation Dashboard

Static interactive dashboard for GitHub Pages, designed from the provided `index(16).html` style pattern and converted from the attached Excel workbooks.

## Included files

```text
root_sourcing_dashboard/
├── index.html
├── data/
│   └── dashboard_data.json
├── tools/
│   └── convert_excel_to_json.py
├── README.md
└── SYSTEM_DESIGN.md
```

## Data sources converted

- `KSN Root Sourcing report App.xlsx`
- `SK Root Sourcing report App.xlsx`
- `Market situation.xlsx`
- `ST_PPDS_Price template.xlsx`

Generated dataset summary:

- Sourcing transaction records: 14,875 rows
- KPI monthly rows: 96 rows
- Volume contribution rows: 96 rows
- Vendor name master rows: 1,694 rows
- Workbook sheets available in Sheet Explorer: 75 sheets

## Main dashboard tabs

1. **Overview** — KPI cards, monthly volume trend, KPI trend, vendor type contribution, top provinces, top vendors.
2. **Sourcing Details** — farmer/broker/regular volume trend, area insights, searchable transaction table.
3. **Market Situation** — tapioca starch/chip/ethanol/dry pulp price trends and Thailand tapioca production chart.
4. **Sheet Explorer** — view and export every sheet from every workbook.
5. **Data Quality** — field completeness, workbook/sheet coverage, data notes.
6. **GitHub Setup** — deployment steps embedded in the dashboard.

## Deploy to GitHub Pages

1. Create a new GitHub repository, for example `root-sourcing-dashboard`.
2. Upload all files in this folder to the repository root.
3. Go to **Settings → Pages**.
4. Select **Deploy from a branch**.
5. Select branch **main** and folder **/root**.
6. Open the generated GitHub Pages URL.

## Update data later

Put the new Excel files in the same project root using these exact names:

```text
KSN Root Sourcing report App.xlsx
SK Root Sourcing report App.xlsx
Market situation.xlsx
ST_PPDS_Price template.xlsx
```

Then run:

```bash
python tools/convert_excel_to_json.py
```

If the Excel files are stored in another folder:

```bash
python tools/convert_excel_to_json.py /path/to/excel-folder
```

Commit the updated `data/dashboard_data.json` to GitHub.

## Notes

- This dashboard is a static web app. It does not require a backend server or database.
- GitHub Pages can serve the JSON file directly.
- Opening `index.html` by double-clicking may block local JSON loading in some browsers. Use GitHub Pages or a local web server such as `python -m http.server`.
- Weighted Starch = `sum(Weigh starch) / sum(Volume kg)`.
- Weighted PPDS = `sum(Amount) / sum(Weigh starch)`.
- Average Price = `sum(Amount) / sum(Volume kg)`.
