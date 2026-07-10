from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

input_folder  = Path(_mp.main_path)
output_folder = input_folder / 'output' / 'json'

"""
extract_03_eut_basic_and_components.py
======================================
Reads 'eut_basic_and_cert_components.xlsx' from input_folder.

Sheets expected:
    'Basico'      → column A = key, column B = value
                    saved as extracted_xlsx_eut_basic.json

    'Componentes' → column A = component name, column B = status string
                    status is mapped to { "exists": bool, "certified": bool }
                    saved as extracted_xlsx_eut_cert_components.json

Component status strings (case-insensitive):
    "NO existe"               → { "exists": false, "certified": false }
    "Existe y certificado"    → { "exists": true,  "certified": true  }
    "Existe y NO certificado" → { "exists": true,  "certified": false }

Install dependency (once):
    pip install openpyxl
"""

import json

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl")
    sys.exit(1)


def _parse_component_value(raw: str) -> dict:
    val = str(raw).strip().lower()
    if 'no existe' in val:
        return {'exists': False, 'certified': False}
    if 'existe y no certificado' in val:
        return {'exists': True, 'certified': False}
    if 'existe y certificado' in val:
        return {'exists': True, 'certified': True}
    if val in ('existe', 'si', 'sí', 'yes'):
        return {'exists': True, 'certified': False}
    return {'exists': False, 'certified': False}


def _sheet_to_kv(ws) -> dict:
    """Read two-column sheet → plain key:value dict, skipping empty/header rows."""
    result = {}
    for key, val in ws.iter_rows(min_col=1, max_col=2, values_only=True):
        k = str(key).strip() if key is not None else ''
        v = str(val).strip() if val is not None else ''
        if not k or k.lower() == 'parametro':
            continue
        result[k] = v
    return result


def extract(xlsx_path: str | Path) -> tuple[dict, dict]:
    """
    Parse xlsx and return (eut_basic, cert_components).

    eut_basic       : { "Class": "Clase II", ... }
    cert_components : { "transformador": {"exists": True, "certified": True}, ... }
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    ws_basic = wb['Basico']
    eut_basic = _sheet_to_kv(ws_basic)

    ws_comp = wb['Componentes']
    raw_comp = _sheet_to_kv(ws_comp)
    cert_components = {k.lower().rstrip(): _parse_component_value(v)
                       for k, v in raw_comp.items()}

    return eut_basic, cert_components


if __name__ == '__main__':

    xlsx_name = 'eut_basic_and_cert_components.xlsx'
    xlsx_file = input_folder / xlsx_name

    if not xlsx_file.is_file():
        print(f"ERROR: '{xlsx_name}' not found in: {input_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)
    print(f"\nFile: {xlsx_file.name}\n")

    eut_basic, cert_components = extract(xlsx_file)

    # ── Save extracted_xlsx_eut_basic.json ────────────────────────────────────
    print("EUT basic:")
    for k, v in eut_basic.items():
        print(f"  {k:<10} {v}")

    eut_file = output_folder / 'extracted_xlsx_eut_basic.json'
    with open(eut_file, 'w', encoding='utf-8') as f:
        json.dump(eut_basic, f, ensure_ascii=False, indent=2)
    print(f"\n  Saved: {eut_file}")

    # ── Save extracted_xlsx_eut_cert_components.json ──────────────────────────
    print("\nComponentes:")
    for k, v in cert_components.items():
        status = 'certificado' if v['certified'] else \
                 ('existe'    if v['exists']    else 'NO existe')
        print(f"  {k:<40} {status}")

    comp_file = output_folder / 'extracted_xlsx_eut_cert_components.json'
    with open(comp_file, 'w', encoding='utf-8') as f:
        json.dump(cert_components, f, ensure_ascii=False, indent=2)
    print(f"\n  Saved: {comp_file}\n")
