#!/usr/bin/env bash
set -euo pipefail
python3 tools/convert_excel_to_json.py
python3 tools/validate_data.py
python3 -m http.server 8000
