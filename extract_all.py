from pathlib import Path


input_folder  = Path(r"C:\Users\nicol\OneDrive\Desktop\MASTER Prueba AI Claude IEC 62368\Prueba")
output_folder = input_folder / 'output'

"""
extract_all.py
==============
Extracts all input data in one run and saves four JSON files:

    sistema.json        ← from *sistema*.htm
    certificadora.json  ← from *certificadora*.pdf
    EUT_basic.json      ← from *eut*/*component*.xlsx  (first group)
    Componentes.json    ← from *eut*/*component*.xlsx  (second group)

Usage:
    python extract_all.py

    Edit the two path variables in the Configuration block below:
        input_folder   — folder containing the .htm, .pdf and .xlsx files
        output_folder  — folder where the four .json files are written

Install dependencies (once):
    pip install pdfplumber openpyxl
"""

import re
import sys
import json
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run:  pip install pdfplumber")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl")
    sys.exit(1)


# ═══════════════════════════════════════════════════════════════════════════════
# 1 — sistema.htm
# ═══════════════════════════════════════════════════════════════════════════════

def extract_sistema(htm_path: str | Path) -> dict:
    """Extract Número de Informe, Fecha de Alta, Cliente, Dirección."""
    text = Path(htm_path).read_text(encoding='utf-8', errors='ignore')

    m = re.search(r'#\d+\s*/\s*([A-Z]{2,4}-\d{2}-\d{2}-\d+)', text)
    informe = m.group(1).strip() if m else ''

    m = re.search(r'#\d+\s*/\s*[A-Z]{2,4}-\d{2}-\d{2}-\d+\s+\t([\d/]+)', text)
    fecha = m.group(1).strip() if m else ''

    m = re.search(r'\n[a-z]+\s+\t([A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ0-9\s\.\,]+?)\s+<http', text)
    cliente = m.group(1).strip() if m else ''

    block_m = re.search(
        r'<http://sistema\.lenor\.com\.ar/Entidad/Edit/\d+>\s*\n'
        r'\t([^\n]+)\n'
        r'([^\n]+)\n',
        text)
    direccion = block_m.group(2).strip() if block_m else ''

    if not direccion:
        m = re.search(
            r'\n([A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ0-9\s]+\d+[^\n]{5,80}'
            r'(?:Argentina|Buenos Aires)[^\n]{0,40})\n',
            text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            if 'fraga' not in candidate.lower():
                direccion = candidate

    return {
        'Número de Informe': informe,
        'Fecha de Alta':     fecha,
        'Cliente':           cliente,
        'Dirección':         direccion,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 2 — certificadora.pdf
# ═══════════════════════════════════════════════════════════════════════════════

def extract_certificadora(pdf_path: str | Path) -> dict:
    """Extract Certificadora, Solicitud, Identificación, Descripción, Marca,
    Modelo, Característica(s)."""
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                rows.extend(table)

    def find(key: str, col: int = 1) -> str:
        for row in rows:
            if row and str(row[0] or '').strip() == key:
                val = row[col] if len(row) > col else ''
                return str(val or '').replace('\n', ' ').strip()
        return ''

    solicitud = find('Número(s) de proceso', 1)

    identificacion = ''
    marca = modelo = caracteristicas = ''
    for i, row in enumerate(rows):
        if row and str(row[0] or '').strip() == 'Etiqueta N°':
            if i + 1 < len(rows):
                identificacion  = str(rows[i+1][0] or '').strip()
                data_row        = rows[i+1]
                marca           = str(data_row[3] or '').strip()           if len(data_row) > 3 else ''
                modelo          = str(data_row[4] or '').strip()           if len(data_row) > 4 else ''
                caracteristicas = str(data_row[5] or '').replace('\n',' ').strip() if len(data_row) > 5 else ''
            break

    descripcion = ''
    for row in rows:
        if row and str(row[0] or '').strip() == 'Producto:':
            for col in range(1, len(row)):
                val = str(row[col] or '').strip()
                if val:
                    descripcion = val
                    break
            break

    return {
        'Certificadora':                  'Lenor OCP',
        'Solicitud':                      solicitud,
        'Identificación':                 identificacion,
        'Descripción del ítem de ensayo': descripcion,
        'Registrada':                     marca,
        'Modelo y/o referencia tipo ':    modelo,
        'Característica(s)':              caracteristicas,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 3 — EUT + components xlsx
# ═══════════════════════════════════════════════════════════════════════════════

def _parse_component(raw: str) -> dict:
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


def extract_eut_components(xlsx_path: str | Path) -> tuple[dict, dict]:
    """Return (eut_data, components) from the combined xlsx."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    eut_data   = {}
    components = {}
    group      = 0

    for key, val in ws.iter_rows(min_col=1, max_col=2, values_only=True):
        k = str(key).strip() if key is not None else ''
        v = str(val).strip() if val is not None else ''
        if not k and not v:
            continue
        if k.lower() == 'key' and v.lower() == 'value':
            group += 1
            continue
        if group == 1 and k and v:
            eut_data[k] = v
        elif group == 2 and k:
            components[k.lower().rstrip()] = _parse_component(v)

    return eut_data, components


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _find(folder: Path, suffix: str, keyword: str) -> Path | None:
    """Find first file in folder matching suffix and keyword (case-insensitive)."""
    for f in folder.iterdir():
        if f.is_file() and f.suffix.lower() == suffix \
                and keyword.lower() in f.name.lower():
            return f
    return None

def _save(data: dict, path: Path):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  ✔ Saved: {path.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder containing the .htm, .pdf and .xlsx source files
    # input_folder  = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where the four .json files will be written
    # output_folder = Path(r'C:\Users\yourname\Documents\job_folder')
    # ── End of configuration ──────────────────────────────────────────────────

    output_folder.mkdir(parents=True, exist_ok=True)
    errors = []
    output = {}   # single combined JSON

    print(f"\n{'─'*55}")
    print(f"  extract_all  —  {input_folder.name}")
    print(f"{'─'*55}\n")

    # ── sistema ───────────────────────────────────────────────────────────────
    htm_file = _find(input_folder, '.htm', 'sistema') or \
               _find(input_folder, '.html', 'sistema')
    if htm_file:
        print(f"sistema  ←  {htm_file.name}")
        sis = extract_sistema(htm_file)
        for k, v in sis.items(): print(f"  {k:<30} {v}")
        output['Sistema'] = sis
    else:
        errors.append("No *sistema*.htm found")
        print("  ⚠ No *sistema*.htm found — skipping")

    # ── certificadora ─────────────────────────────────────────────────────────
    print()
    pdf_file = _find(input_folder, '.pdf', 'certificadora')
    if pdf_file:
        print(f"certificadora  ←  {pdf_file.name}")
        cert = extract_certificadora(pdf_file)
        for k, v in cert.items(): print(f"  {k:<35} {v}")
        output['Certificadora'] = cert
    else:
        errors.append("No *certificadora*.pdf found")
        print("  ⚠ No *certificadora*.pdf found — skipping")

    # ── EUT + components ──────────────────────────────────────────────────────
    print()
    xlsx_file = _find(input_folder, '.xlsx', 'eut') or \
                _find(input_folder, '.xlsx', 'component')
    if xlsx_file:
        print(f"EUT + Componentes  ←  {xlsx_file.name}")
        eut, comp = extract_eut_components(xlsx_file)

        print("  EUT basic:")
        for k, v in eut.items(): print(f"    {k:<10} {v}")
        output['EUT_basic'] = eut

        print("  Componentes:")
        for k, v in comp.items():
            status = 'certificado' if v['certified'] else \
                     ('existe' if v['exists'] else 'NO existe')
            print(f"    {k:<40} {status}")
        output['Componentes'] = comp
    else:
        errors.append("No *eut*/*component*.xlsx found")
        print("  ⚠ No *eut*/*component*.xlsx found — skipping")

    # ── Save single combined JSON ──────────────────────────────────────────────
    out_file = output_folder / 'extracted_data.json'
    _save(output, out_file)

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{'─'*55}")
    if errors:
        print(f"  Done with {len(errors)} warning(s):")
        for e in errors: print(f"    ⚠ {e}")
    else:
        print("  All extractions completed successfully.")
    print(f"{'─'*55}\n")
