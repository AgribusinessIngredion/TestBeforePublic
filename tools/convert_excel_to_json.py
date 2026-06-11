#!/usr/bin/env python3
"""Convert Ingredion root sourcing Excel workbooks into a compact JSON dataset for GitHub Pages.
Uses only Python standard libraries to read .xlsx OpenXML contents."""
from __future__ import annotations
import json, math, os, posixpath, re, sys, zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Dict, List, Tuple

NS = {
    'main': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'rel': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pkgrel': 'http://schemas.openxmlformats.org/package/2006/relationships',
}
MONTH_ORDER = {m:i+1 for i,m in enumerate(['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'])}

SOURCE_FILES = [
    ('KSN Root Sourcing report App.xlsx', 'Kalasin'),
    ('SK Root Sourcing report App.xlsx', 'Sikhiu'),
    ('Market situation.xlsx', 'Market'),
    ('ST_PPDS_Price template.xlsx', 'Template'),
]


def col_to_idx(ref: str) -> int:
    m = re.match(r'([A-Z]+)', ref or 'A1')
    if not m:
        return 0
    n = 0
    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64
    return n - 1


def excel_date_to_iso(value: Any) -> Any:
    if isinstance(value, (int, float)) and 20000 <= float(value) <= 70000:
        # Excel 1900 date system with leap-year bug adjustment.
        dt = datetime(1899, 12, 30) + timedelta(days=float(value))
        return dt.date().isoformat()
    return value


def clean_number(x: Any) -> Any:
    if isinstance(x, float):
        if math.isnan(x) or math.isinf(x):
            return None
        return round(x, 10)
    return x


def trim_row(row: List[Any]) -> List[Any]:
    while row and (row[-1] is None or row[-1] == ''):
        row.pop()
    return row


def parse_shared(z: zipfile.ZipFile) -> List[str]:
    try:
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    out: List[str] = []
    for si in root.findall('main:si', NS):
        texts = [t.text or '' for t in si.findall('.//main:t', NS)]
        out.append(''.join(texts))
    return out


def parse_workbook(path: Path) -> Dict[str, List[List[Any]]]:
    with zipfile.ZipFile(path) as z:
        shared = parse_shared(z)
        wbroot = ET.fromstring(z.read('xl/workbook.xml'))
        relroot = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rels = {r.attrib['Id']: r.attrib['Target'] for r in relroot}
        sheets: List[Tuple[str, str]] = []
        for sh in wbroot.find('main:sheets', NS):
            rid = sh.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            target = rels[rid]
            if not target.startswith('/'):
                target = posixpath.normpath(posixpath.join('xl', target))
            else:
                target = target.lstrip('/')
            sheets.append((sh.attrib['name'], target))

        data: Dict[str, List[List[Any]]] = {}
        for sheet_name, target in sheets:
            root = ET.fromstring(z.read(target))
            sheetData = root.find('main:sheetData', NS)
            rows_sparse: List[Dict[int, Any]] = []
            max_col = -1
            if sheetData is None:
                data[sheet_name] = []
                continue
            for row_el in sheetData.findall('main:row', NS):
                vals: Dict[int, Any] = {}
                for c in row_el.findall('main:c', NS):
                    ci = col_to_idx(c.attrib.get('r', 'A1'))
                    t = c.attrib.get('t')
                    v = c.find('main:v', NS)
                    is_el = c.find('main:is', NS)
                    val = None
                    if t == 'inlineStr' and is_el is not None:
                        val = ''.join([(tt.text or '') for tt in is_el.findall('.//main:t', NS)])
                    elif v is not None and v.text is not None:
                        txt = v.text
                        if t == 's':
                            try:
                                val = shared[int(txt)]
                            except Exception:
                                val = txt
                        elif t == 'b':
                            val = txt == '1'
                        else:
                            try:
                                f = float(txt)
                                val = int(f) if f.is_integer() else f
                            except ValueError:
                                val = txt
                    if val is not None and val != '':
                        vals[ci] = clean_number(val)
                        if ci > max_col:
                            max_col = ci
                rows_sparse.append(vals)
            while rows_sparse and not rows_sparse[-1]:
                rows_sparse.pop()
            matrix: List[List[Any]] = []
            for vals in rows_sparse:
                row = [vals.get(i) for i in range(max_col + 1)] if max_col >= 0 else []
                matrix.append(trim_row(row))
            # Drop fully blank leading rows to keep tables compact.
            while matrix and not matrix[0]:
                matrix.pop(0)
            data[sheet_name] = matrix
        return data


def find_header_row(rows: List[List[Any]], required: List[str]) -> int:
    req = [r.lower() for r in required]
    for i, row in enumerate(rows[:15]):
        lows = [str(x).strip().lower() for x in row if x is not None]
        if all(any(r in cell for cell in lows) for r in req):
            return i
    return 0


def row_to_objects(rows: List[List[Any]], header_idx: int = 0) -> List[Dict[str, Any]]:
    if not rows or header_idx >= len(rows):
        return []
    headers = [str(h).strip() if h is not None else f'Column {i+1}' for i, h in enumerate(rows[header_idx])]
    objects: List[Dict[str, Any]] = []
    for row in rows[header_idx+1:]:
        if not any(v is not None and v != '' for v in row):
            continue
        obj = {}
        for i, h in enumerate(headers):
            if h == '':
                h = f'Column {i+1}'
            obj[h] = clean_number(row[i] if i < len(row) else None)
        objects.append(obj)
    return objects


def get_any(obj: Dict[str, Any], *names: str) -> Any:
    norm = {k.lower().replace('.', '').replace(' ', ''): k for k in obj}
    for name in names:
        key = name.lower().replace('.', '').replace(' ', '')
        if key in norm:
            return obj.get(norm[key])
    return None


def normalize_records(rows: List[List[Any]], plant: str, workbook: str, sheet: str) -> List[Dict[str, Any]]:
    header_idx = find_header_row(rows, ['year', 'month', 'vendor'])
    objs = row_to_objects(rows, header_idx)
    out = []
    for o in objs:
        year = get_any(o, 'Year')
        month = get_any(o, 'Month')
        vendor = get_any(o, 'Vendor')
        if year is None or month is None or not vendor:
            continue
        volkg = get_any(o, 'Volume, kg', 'Volume, kg.')
        volmt = get_any(o, 'Volume, MT')
        if volmt is None and isinstance(volkg, (int, float)):
            volmt = volkg / 1000
        out.append({
            'plant': plant,
            'workbook': workbook,
            'sheet': sheet,
            'year': int(year) if isinstance(year, (int, float)) else year,
            'month': str(month),
            'monthNo': MONTH_ORDER.get(str(month), 0),
            'rootsType': get_any(o, 'Roots type'),
            'vendorType': get_any(o, 'Vendor type'),
            'vendor': str(vendor).strip(),
            'district': get_any(o, 'District'),
            'province': get_any(o, 'Province'),
            'volumeKg': clean_number(volkg),
            'volumeMt': clean_number(volmt),
            'starchContent': clean_number(get_any(o, 'Starch content')),
            'weighStarch': clean_number(get_any(o, 'Weigh starch')),
            'price': clean_number(get_any(o, 'Price')),
            'amount': clean_number(get_any(o, 'Amount')),
            'ppds': clean_number(get_any(o, 'PPDS')),
        })
    return out


def normalize_kpi(rows: List[List[Any]], workbook: str, sheet: str) -> List[Dict[str, Any]]:
    objs = row_to_objects(rows, find_header_row(rows, ['year', 'month', 'plant']))
    out = []
    for o in objs:
        if get_any(o, 'Year') is None or get_any(o, 'Month') is None:
            continue
        plant_raw = get_any(o, 'Plant')
        out.append({
            'workbook': workbook,
            'sheet': sheet,
            'year': int(get_any(o, 'Year')) if isinstance(get_any(o, 'Year'), (int,float)) else get_any(o, 'Year'),
            'month': str(get_any(o, 'Month')),
            'monthNo': MONTH_ORDER.get(str(get_any(o, 'Month')), 0),
            'plant': 'Kalasin' if str(plant_raw).lower().startswith('kalasin') else ('Sikhiu' if str(plant_raw).lower().startswith('sikhiu') else plant_raw),
            'le0Ppds': clean_number(get_any(o, 'LE0 PPDS, THB/kg')),
            'lvPpds': clean_number(get_any(o, 'LV PPDS, THB/kg')),
            'actPpds': clean_number(get_any(o, 'Act. PPDS, THB/kg')),
            'ttsaPrice': clean_number(get_any(o, 'TTSA Price')),
            'actPpdsCostAgainstTtsa': clean_number(get_any(o, 'Act. PPDS cost against TTSA, THB/kg')),
            'estPpdsAccuracy': clean_number(get_any(o, 'Est. PPDS accuracy, %')),
            'actPpdsAccuracy': clean_number(get_any(o, 'Act. PPDS accuracy, %')),
            'le0Starch': clean_number(get_any(o, 'LE0 Starch content, %')),
            'lvStarch': clean_number(get_any(o, 'LV Starch content, %')),
            'actStarch': clean_number(get_any(o, 'Act. Starch content, %')),
            'le0Price': clean_number(get_any(o, 'LE0 Price THB/kg')),
            'lvPrice': clean_number(get_any(o, 'LV Price THB/kg')),
            'actPrice': clean_number(get_any(o, 'Act. Price THB/kg')),
            'estFarmerContribution': clean_number(get_any(o, 'Est. Farmer contribution , %')),
            'estImpurity': clean_number(get_any(o, 'Est. impurity, %')),
            'actImpurity': clean_number(get_any(o, 'Act. impurity, %')),
        })
    return out


def normalize_volume(rows: List[List[Any]], plant: str, workbook: str, sheet: str) -> List[Dict[str, Any]]:
    objs = row_to_objects(rows, find_header_row(rows, ['year', 'month']))
    out = []
    for o in objs:
        if get_any(o, 'Year') is None or get_any(o, 'Month') is None:
            continue
        out.append({
            'workbook': workbook,
            'sheet': sheet,
            'plant': plant,
            'year': int(get_any(o, 'Year')) if isinstance(get_any(o, 'Year'), (int,float)) else get_any(o, 'Year'),
            'month': str(get_any(o, 'Month')),
            'monthNo': MONTH_ORDER.get(str(get_any(o, 'Month')), 0),
            'brokerVolume': clean_number(get_any(o, 'Broker volume, tons')),
            'farmerVolume': clean_number(get_any(o, 'Farmer volume, tons')),
            'regularVolume': clean_number(get_any(o, 'Regular volume, tons')),
        })
    return out


def convert_date_columns(rows: List[List[Any]]) -> List[List[Any]]:
    if not rows:
        return rows
    date_cols = set()
    for ri in range(min(5, len(rows))):
        for ci, val in enumerate(rows[ri]):
            if isinstance(val, str) and 'date' in val.lower():
                date_cols.add(ci)
    if not date_cols:
        return rows
    new_rows = []
    for r in rows:
        nr = []
        for ci, val in enumerate(r):
            nr.append(excel_date_to_iso(val) if ci in date_cols else val)
        new_rows.append(nr)
    return new_rows


def extract_market(workbooks: Dict[str, Dict[str, List[List[Any]]]]) -> Dict[str, Any]:
    market_wb = workbooks.get('Market situation.xlsx', {})
    def rows(sheet): return convert_date_columns(market_wb.get(sheet, []))

    # Starch price
    starch = []
    sp = rows('Tapioca Starch Price_1')
    for r in sp[1:]:
        if len(r) >= 2 and r[0] and isinstance(r[1], (int, float)):
            starch.append({'date': r[0], 'domesticPrice': r[1], 'fobBangkok': r[3] if len(r)>3 else None, 'fobHcmLowest': r[4] if len(r)>4 else None, 'fobHcmHighest': r[5] if len(r)>5 else None})

    chip = []
    cp = rows('Tapioca Chip Price_1')
    for r in cp[1:]:
        if len(r) >= 2 and r[0]:
            chip.append({'date': r[0], 'thbKg': r[1] if len(r)>1 else None, 'thSrichangUsdMt': r[2] if len(r)>2 else None, 'vnChinaUsdMt': r[3] if len(r)>3 else None, 'vnKoreaUsdMt': r[4] if len(r)>4 else None})

    ethanol = []
    et = rows('Ethanol Price_1')
    # Try to locate row with Month/THB./Liter
    hi = find_header_row(et, ['month'])
    for r in et[hi+1:]:
        if len(r) >= 3 and r[1] is not None and isinstance(r[2], (int, float)):
            ethanol.append({'month': r[1], 'thbLiter': r[2]})

    dry_monthly = []
    dm = rows('Dry Pulp Monthly price_1')
    for r in dm[1:]:
        if len(r) >= 2 and r[0] is not None and isinstance(r[1], (int, float)):
            dry_monthly.append({'month': r[0], 'price': r[1]})

    dry_daily = []
    dd = rows('Dry Pulp Daily price_1')
    for r in dd[1:]:
        if len(r) >= 4 and r[0] is not None:
            dry_daily.append({'date': r[0], 'start': r[1] if len(r)>1 else None, 'end': r[2] if len(r)>2 else None, 'average': r[3] if len(r)>3 else None})

    production = []
    prod = market_wb.get('Tapioca Production_1', [])
    for r in prod[4:]:
        if len(r) >= 6 and r[1] is not None:
            production.append({'cropYear': r[1], 'yieldTonRai': r[2] if len(r)>2 else None, 'plantAreaRai': r[4] if len(r)>4 else None, 'productionTon': r[5] if len(r)>5 else None})

    return {'starchPrice': starch, 'chipPrice': chip, 'ethanolPrice': ethanol, 'dryPulpMonthly': dry_monthly, 'dryPulpDaily': dry_daily, 'production': production}


def sheet_summary(rows: List[List[Any]]) -> Dict[str, Any]:
    max_cols = max((len(r) for r in rows), default=0)
    non_empty = sum(1 for r in rows for v in r if v is not None and v != '')
    return {'rows': len(rows), 'columns': max_cols, 'nonEmptyCells': non_empty}


def build_dataset(base_dir: Path, out_path: Path) -> Dict[str, Any]:
    workbooks: Dict[str, Dict[str, List[List[Any]]]] = {}
    workbook_sheets = []
    records: List[Dict[str, Any]] = []
    kpis: List[Dict[str, Any]] = []
    volumes: List[Dict[str, Any]] = []
    name_rows: List[Dict[str, Any]] = []

    for filename, plant in SOURCE_FILES:
        path = base_dir / filename
        if not path.exists():
            continue
        wbdata = parse_workbook(path)
        workbooks[filename] = wbdata
        for sheet, rows in wbdata.items():
            visible_rows = convert_date_columns(rows)
            workbook_sheets.append({'workbook': filename, 'sheet': sheet, 'summary': sheet_summary(visible_rows), 'rows': visible_rows})
            if filename.startswith('SK') and sheet == 'RegRoot.SK':
                records.extend(normalize_records(rows, 'Sikhiu', filename, sheet))
            elif filename.startswith('KSN') and sheet == 'RegRoot.KSN':
                records.extend(normalize_records(rows, 'Kalasin', filename, sheet))
            elif sheet.startswith('RawKPI'):
                kpis.extend(normalize_kpi(rows, filename, sheet))
            elif sheet in ('Volume.KSN', 'VolFarmerContri.SK'):
                volumes.extend(normalize_volume(rows, 'Kalasin' if 'KSN' in sheet else 'Sikhiu', filename, sheet))
            elif sheet.startswith('Name.'):
                objs = row_to_objects(rows, find_header_row(rows, ['vendor']))
                for o in objs:
                    o['plant'] = 'Kalasin' if 'KSN' in sheet else 'Sikhiu'
                    name_rows.append(o)

    market = extract_market(workbooks)
    dataset = {
        'metadata': {
            'title': 'Root Sourcing & Market Situation Dashboard',
            'generatedAt': datetime.now(timezone.utc).isoformat(),
            'sourceFiles': [{'file': f, 'category': p} for f,p in SOURCE_FILES if (base_dir/f).exists()],
            'counts': {'sourcingRecords': len(records), 'kpiRows': len(kpis), 'volumeRows': len(volumes), 'nameRows': len(name_rows), 'workbookSheets': len(workbook_sheets)}
        },
        'sourcing': {'records': records, 'kpis': kpis, 'volumes': volumes, 'names': name_rows},
        'market': market,
        'workbookSheets': workbook_sheets,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(dataset, ensure_ascii=False, separators=(',', ':')), encoding='utf-8')
    return dataset

# -----------------------------------------------------------------------------
# Production helpers: CSV exports, manifest, and data-quality reports
# -----------------------------------------------------------------------------
def safe_filename(value: Any) -> str:
    text = str(value or 'sheet').strip()
    text = re.sub(r'[\\/:*?"<>|]+', '_', text)
    text = re.sub(r'\s+', '_', text)
    return (text[:120] or 'sheet')


def collect_headers(rows: List[Dict[str, Any]]) -> List[str]:
    headers: List[str] = []
    seen = set()
    for row in rows:
        for key in row.keys():
            if key not in seen:
                headers.append(key)
                seen.add(key)
    return headers


def write_dict_csv(path: Path, rows: List[Dict[str, Any]], headers: List[str] | None = None) -> None:
    import csv
    path.parent.mkdir(parents=True, exist_ok=True)
    if headers is None:
        headers = collect_headers(rows) if rows else []
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
        writer.writeheader()
        for row in rows:
            writer.writerow({h: '' if row.get(h) is None else row.get(h) for h in headers})


def write_matrix_csv(path: Path, rows: List[List[Any]]) -> None:
    import csv
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        for row in rows:
            writer.writerow(['' if v is None else v for v in row])


def export_csv_files(dataset: Dict[str, Any], project_root: Path) -> List[Dict[str, Any]]:
    csv_dir = project_root / 'data' / 'csv'
    records = dataset.get('sourcing', {}).get('records', [])
    kpis = dataset.get('sourcing', {}).get('kpis', [])
    volumes = dataset.get('sourcing', {}).get('volumes', [])
    names = dataset.get('sourcing', {}).get('names', [])

    write_dict_csv(csv_dir / 'sourcing_records_all.csv', records)
    write_dict_csv(csv_dir / 'sk_regroot.csv', [r for r in records if r.get('plant') == 'Sikhiu'])
    write_dict_csv(csv_dir / 'ksn_regroot.csv', [r for r in records if r.get('plant') == 'Kalasin'])
    write_dict_csv(csv_dir / 'kpi_monthly_all.csv', kpis)
    write_dict_csv(csv_dir / 'sk_kpi_monthly.csv', [r for r in kpis if r.get('plant') == 'Sikhiu'])
    write_dict_csv(csv_dir / 'ksn_kpi_monthly.csv', [r for r in kpis if r.get('plant') == 'Kalasin'])
    write_dict_csv(csv_dir / 'volume_contribution_all.csv', volumes)
    write_dict_csv(csv_dir / 'sk_volume_contribution.csv', [r for r in volumes if r.get('plant') == 'Sikhiu'])
    write_dict_csv(csv_dir / 'ksn_volume_contribution.csv', [r for r in volumes if r.get('plant') == 'Kalasin'])
    write_dict_csv(csv_dir / 'vendor_master_all.csv', names)

    for market_key, rows in dataset.get('market', {}).items():
        if isinstance(rows, list):
            write_dict_csv(csv_dir / f'market_{safe_filename(market_key)}.csv', rows)

    sheet_index: List[Dict[str, Any]] = []
    for item in dataset.get('workbookSheets', []):
        workbook = safe_filename(item.get('workbook', 'workbook'))
        sheet = safe_filename(item.get('sheet', 'sheet'))
        rel_path = f'data/csv/by_sheet/{workbook}/{sheet}.csv'
        write_matrix_csv(project_root / rel_path, item.get('rows', []))
        sheet_index.append({
            'workbook': item.get('workbook'),
            'sheet': item.get('sheet'),
            'csvPath': rel_path,
            **item.get('summary', {})
        })
    write_dict_csv(csv_dir / 'sheet_index.csv', sheet_index, ['workbook', 'sheet', 'csvPath', 'rows', 'columns', 'nonEmptyCells'])

    return sheet_index


def build_quality_report(dataset: Dict[str, Any]) -> Dict[str, Any]:
    from collections import Counter
    records = dataset.get('sourcing', {}).get('records', [])
    missing = [r for r in records if r.get('ppds') in (None, '')]
    zero = [r for r in records if r.get('ppds') == 0]
    critical = [
        r for r in missing
        if (r.get('volumeMt') or 0) > 0 or (r.get('amount') or 0) > 0 or (r.get('price') or 0) > 0
    ]
    by_source = Counter((r.get('plant'), r.get('workbook'), r.get('sheet')) for r in missing)
    by_month = Counter((r.get('plant'), r.get('year'), r.get('month')) for r in missing)
    total = len(records)
    return {
        'generatedAt': datetime.now(timezone.utc).isoformat(),
        'checks': {
            'ppds': {
                'totalRows': total,
                'completeRows': total - len(missing),
                'missingRows': len(missing),
                'zeroRows': len(zero),
                'criticalMissingRows': len(critical),
                'completePercent': round((total - len(missing)) / total * 100, 2) if total else 0,
                'bySource': [
                    {'plant': k[0], 'workbook': k[1], 'sheet': k[2], 'missingRows': v}
                    for k, v in by_source.most_common()
                ],
                'byMonth': [
                    {'plant': k[0], 'year': k[1], 'month': k[2], 'missingRows': v}
                    for k, v in by_month.most_common()
                ]
            },
            'counts': dataset.get('metadata', {}).get('counts', {})
        },
        'recommendation': 'If criticalMissingRows = 0, missing PPDS rows are likely zero-volume audit records. Keep them but exclude them from average PPDS calculations.'
    }


def export_quality_files(dataset: Dict[str, Any], project_root: Path) -> Dict[str, Any]:
    records = dataset.get('sourcing', {}).get('records', [])
    missing = [r for r in records if r.get('ppds') in (None, '')]
    zero = [r for r in records if r.get('ppds') == 0]
    critical = [
        r for r in missing
        if (r.get('volumeMt') or 0) > 0 or (r.get('amount') or 0) > 0 or (r.get('price') or 0) > 0
    ]
    report = build_quality_report(dataset)
    data_dir = project_root / 'data'
    data_dir.mkdir(parents=True, exist_ok=True)
    (data_dir / 'quality_report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    write_dict_csv(data_dir / 'ppds_missing_records.csv', missing)
    write_dict_csv(data_dir / 'ppds_zero_records.csv', zero)
    write_dict_csv(data_dir / 'critical_missing_ppds_records.csv', critical)
    return report


def export_manifest(dataset: Dict[str, Any], project_root: Path, quality: Dict[str, Any]) -> None:
    data_files = []
    for path in sorted((project_root / 'data').rglob('*')):
        if path.is_file():
            data_files.append({
                'path': str(path.relative_to(project_root)).replace('\\', '/'),
                'sizeBytes': path.stat().st_size
            })
    manifest = {
        'project': 'Root Sourcing & Market Situation Dashboard',
        'generatedAt': datetime.now(timezone.utc).isoformat(),
        'dashboardEntry': 'index.html',
        'mainDataFile': 'data/dashboard_data.json',
        'qualityReport': 'data/quality_report.json',
        'sourceExcelFolder': 'data/source_excel/',
        'csvFolder': 'data/csv/',
        'counts': dataset.get('metadata', {}).get('counts', {}),
        'ppdsQuality': quality.get('checks', {}).get('ppds', {}),
        'files': data_files,
    }
    (project_root / 'data' / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')


def resolve_source_dir(project_root: Path, argv: List[str]) -> Path:
    if len(argv) > 1:
        return Path(argv[1]).resolve()
    preferred = project_root / 'data' / 'source_excel'
    if preferred.exists():
        return preferred
    return project_root


if __name__ == '__main__':
    project_root = Path(__file__).resolve().parents[1]
    source_dir = resolve_source_dir(project_root, sys.argv)
    out = project_root / 'data' / 'dashboard_data.json'
    dataset = build_dataset(source_dir, out)
    export_csv_files(dataset, project_root)
    quality = export_quality_files(dataset, project_root)
    export_manifest(dataset, project_root, quality)
    print(json.dumps({
        'sourceDir': str(source_dir),
        'metadata': dataset.get('metadata', {}),
        'ppdsQuality': quality.get('checks', {}).get('ppds', {})
    }, ensure_ascii=False, indent=2))
