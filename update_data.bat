@echo off
python tools\convert_excel_to_json.py
python tools\validate_data.py
python -m http.server 8000
