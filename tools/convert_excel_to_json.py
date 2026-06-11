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

if __name__ == '__main__':
    project_root = Path(__file__).resolve().parents[1]
    base = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else project_root
    out = project_root / 'data' / 'dashboard_data.json'
    ds = build_dataset(base, out)
    print(json.dumps(ds['metadata'], ensure_ascii=False, indent=2))
