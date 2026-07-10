from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

input_folder = Path(_mp.main_path)
output_file = input_folder / 'output' / 'json' / 'extracted_htm_sistema.json'

Path(input_folder / "output").mkdir(exist_ok=True)
Path(input_folder / "output/json").mkdir(parents=True, exist_ok=True)

"""
extract_sistema.py
==================
Extracts the following fields from a Lenor sistema.htm file:
  - Número de Informe
  - Fecha de Alta
  - Cliente
  - Dirección  (the one immediately after "Nombre del Solicitante", NOT Fraga/Lenor)

Usage:
    python extract_sistema.py 596079_sistema.htm
    python extract_sistema.py 596079_sistema.htm --json   # output as JSON

The extracted values are printed to stdout and also returned as a dict
when imported as a module:
    from extract_sistema import extract
    data = extract("596079_sistema.htm")
"""

import re
import sys
import json


def extract(htm_path: str | Path) -> dict:
    """
    Parse a Lenor sistema.htm and return a dict with:
        Número de Informe, Fecha de Alta, Cliente, Dirección
    """
    text = Path(htm_path).read_text(encoding='utf-8', errors='ignore')

    # ── The key block looks like this in the file: ─────────────────────────
    # #596079 / AJP-04-26-6079     07/04/2026
    # macevedo     NETWORK BROADCAST SA <http://...>
    #     HORACIO ARRIGO
    # PARANA 771 4TO PISO, Buenos Aires, Argentina
    # ───────────────────────────────────────────────────────────────────────

    # Número de Informe  e.g. "AJP-04-26-6079" or "HS -05-26-1095"
    # Tolerate optional whitespace between the letter code and the first dash
    # (the system sometimes renders it as "HS -05-26-1095").
    m = re.search(r'#\d+\s*/\s*([A-Z]{2,4}\s*-\d{2}-\d{2}-\d+)', text)
    informe = re.sub(r'\s+', '', m.group(1)) if m else ''

    # Fecha de Alta  e.g. "07/04/2026"
    # Appears on the same line as the informe number, after a tab
    m = re.search(
        r'#\d+\s*/\s*[A-Z]{2,4}\s*-\d{2}-\d{2}-\d+\s+\t([\d/]+)', text)
    fecha = m.group(1).strip() if m else ''

    # Cliente  e.g. "NETWORK BROADCAST SA"
    # Appears after a username followed by a tab, before a URL. The username
    # may contain dots, dashes or digits (e.g. "macevedo" or "m.villagra").
    m = re.search(
        r'\n[a-z][a-z0-9._-]*\s+\t([A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ0-9\s\.\,]+?)\s+<http',
        text)
    cliente = m.group(1).strip() if m else ''

    # Dirección  — immediately after the contact name line, before "Sector"
    # Pattern: after cliente block → contact name line → address line
    # The address is the line that comes right after the contact person name
    # and ends before "Sector" or the next section.
    # We look for the block between the cliente URL and "Sector":
    #   \tHORACIO ARRIGO\nPARANA 771 4TO PISO, Buenos Aires, Argentina\nSector
    block_m = re.search(
        r'<http://sistema\.lenor\.com\.ar/Entidad/Edit/\d+>\s*\n'  # end of cliente URL
        r'\t([^\n]+)\n'                                              # contact person (skip)
        r'([^\n]+)\n',                                               # ← address line
        text)
    direccion = block_m.group(2).strip() if block_m else ''

    # Fallback: search for the address pattern directly
    # (address looks like "SOMETHING NNN, City, Country")
    if not direccion:
        m = re.search(
            r'\n([A-ZÁÉÍÓÚ][A-ZÁÉÍÓÚ0-9\s]+\d+[^\n]{5,80}'
            r'(?:Argentina|Buenos Aires)[^\n]{0,40})\n',
            text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            # Exclude Lenor's own address (contains "Fraga" or "fraga")
            if 'fraga' not in candidate.lower():
                direccion = candidate

    return {
        'Número de Informe': informe,
        'Fecha de Alta':     fecha,
        'Cliente':           cliente,
        'Dirección':         direccion,
    }


if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder where the sistema file is located (.txt, .htm or .html)
    # input_folder = Path(r'C:\Users\yourname\Documents\job_folder')
    # ── End of configuration ──────────────────────────────────────────────────

    # Find the sistema file: accepts .txt, .htm or .html
    candidates = [f for f in input_folder.iterdir()
                  if f.is_file()
                  and f.suffix.lower() in ('.txt', '.htm', '.html')
                  and 'sistema' in f.name.lower()]

    if not candidates:
        print(f"ERROR: No sistema .txt/.htm file found in: {input_folder}")
        sys.exit(1)

    htm_file    = candidates[0]
    # output_file = input_folder / 'sistema.json'

    print(f"\nFile: {htm_file.name}\n")

    data = extract(htm_file)

    print()
    for key, val in data.items():
        print(f"  {key:<30} {val}")

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n  [OK] Saved: {output_file}\n")