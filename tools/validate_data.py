#!/usr/bin/env python3
"""Validate generated dashboard data before publishing to GitHub Pages.

Default behavior:
- Reads data/dashboard_data.json
- Writes data/quality_report.json and PPDS issue CSV files
- Exits with code 1 only if critical PPDS missing rows are found.
"""
from __future__ import annotations
import csv, json, sys
from pathlib import Path
from datetime import datetime, timezone
from collections import Counter
from typing import Any, Dict, List


def collect_headers(rows: List[Dict[str, Any]]) -> List[str]:
    headers, seen = [], set()
    for row in rows:
        for key in row:
            if key not in seen:
                headers.append(key); seen.add(key)
    return headers


def write_csv(path: Path, rows: List[Dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = collect_headers(rows) if rows else []
    with path.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
        writer.writeheader()
        for row in rows:
            writer.writerow({h: '' if row.get(h) is None else row.get(h) for h in headers})


def build_report(dataset: Dict[str, Any]) -> tuple[Dict[str, Any], List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, Any]]]:
    records = dataset.get('sourcing', {}).get('records', [])
    missing = [r for r in records if r.get('ppds') in (None, '')]
    zero = [r for r in records if r.get('ppds') == 0]
    critical = [r for r in missing if (r.get('volumeMt') or 0) > 0 or (r.get('amount') or 0) > 0 or (r.get('price') or 0) > 0]
    by_source = Counter((r.get('plant'), r.get('workbook'), r.get('sheet')) for r in missing)
    total = len(records)
    report = {
        'generatedAt': datetime.now(timezone.utc).isoformat(),
        'checks': {
            'ppds': {
                'totalRows': total,
                'completeRows': total - len(missing),
                'missingRows': len(missing),
                'zeroRows': len(zero),
                'criticalMissingRows': len(critical),
                'completePercent': round((total - len(missing)) / total * 100, 2) if total else 0,
                'bySource': [{'plant': k[0], 'workbook': k[1], 'sheet': k[2], 'missingRows': v} for k, v in by_source.most_common()],
            },
            'counts': dataset.get('metadata', {}).get('counts', {})
        }
    }
    return report, missing, zero, critical


def main() -> int:
    project_root = Path(__file__).resolve().parents[1]
    data_path = project_root / 'data' / 'dashboard_data.json'
    if not data_path.exists():
        print('ERROR: data/dashboard_data.json not found. Run tools/convert_excel_to_json.py first.', file=sys.stderr)
        return 2
    dataset = json.loads(data_path.read_text(encoding='utf-8'))
    report, missing, zero, critical = build_report(dataset)
    (project_root / 'data' / 'quality_report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    write_csv(project_root / 'data' / 'ppds_missing_records.csv', missing)
    write_csv(project_root / 'data' / 'ppds_zero_records.csv', zero)
    write_csv(project_root / 'data' / 'critical_missing_ppds_records.csv', critical)
    ppds = report['checks']['ppds']
    print(json.dumps(ppds, ensure_ascii=False, indent=2))
    if ppds['criticalMissingRows'] > 0:
        print('ERROR: Critical missing PPDS rows found. Please review data/critical_missing_ppds_records.csv', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
