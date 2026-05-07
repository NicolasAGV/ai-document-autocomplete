from pathlib import Path

input_folder  = Path(r'C:\Users\yourname\Documents\job_folder')
output_folder = Path(r'C:\Users\yourname\Documents\job_folder')


"""
translate_clauses.py
=====================
Reads the clause logic from "New_Prompt_clauses__IEC62368.xlsx" and the
extracted data from "extracted_data.json", evaluates every IF condition
against the actual data, and produces "changes.json" — a ready-to-apply
dictionary that tells modify_docx_table.py exactly what to write in each
row (obs and/or vrd), with no further logic needed.

    Input:
        *clauses*.xlsx      — logic spreadsheet  (Clauses sheet)
        extracted_data.json — data from extract_all.py

    Output:
        changes.json        — { "row_index": {"obs": "...", "vrd": "..."}, ... }

Workflow:
    extract_all.py  →  extracted_data.json  ─┐
    *clauses*.xlsx                           ├─→  translate_clauses.py  →  changes.json
                                              │                               │
                                              └───────────────────────────────┘
                                                                              ↓
                                                              modify_docx_table.py  →  .docx

Usage:
    python translate_clauses.py

    Edit the Configuration block below:
        input_folder   — folder with the .xlsx and extracted_data.json
        output_folder  — folder where changes.json is written

Install dependencies (once):
    pip install openpyxl
"""

import re
import sys
import json

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Run:  pip install openpyxl")
    sys.exit(1)


# ── Row index map: clause id → row index in the clause table ─────────────────
# This is fixed for this .docx template (368_C_II_espigas_inc.docx)
# "again" and "(Nth row below clause)" are handled by offset from the base row.
CLAUSE_ROW = {
    '5.4.4.4':   46,
    '5.4.4.6.2': 50,
    '5.5.2':     86,
    '5.5.2.2':   88,
    '5.5.7':     93,
    '5.7.2.1':   118,
    '6.4.8.2.2': 35,
    'F.3.2.1':   55,
    'F.3.2.2':   56,
    'F.3.3.4':   61,   # first F.3.3.4 row (voltage)
    'F.3.3.4b':  62,   # second F.3.3.4 row (frequency)
    'F.3.3.6':   63,
    'F.4':       84,
    'G.2':       101,
    'G.2.1':     102,
    'G.3.4':     118,
    'G.6':       168,
    'G.6.1':     169,
    'G.8':       193,
    'G.8.1':     194,
    'G.11':      215,
    'G.11.1':    216,
    'G.12':      219,
    'P.1':       353,
    'P.2.2':     354,
    'S.1':       389,
}


# ═══════════════════════════════════════════════════════════════════════════════
# Condition evaluator
# ═══════════════════════════════════════════════════════════════════════════════

def evaluate(condition: str | None,
             data: dict, char: str) -> bool:
    """
    Evaluate an IF condition string against extracted_data.

    data       : full extracted_data dict
    char       : Característica(s) string from Certificadora
    """
    if condition is None:
        return True   # no condition → always apply

    c = condition.strip().lower()

    eut   = data.get('EUT_basic', {})
    comps = data.get('Componentes', {})

    # ── component exist ───────────────────────────────────────────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"\s+(?:exist|exists)\s*$', c)
    if m:
        name = m.group(1).lower()
        return comps.get(name, {}).get('exists', False)

    # ── component OR exist ────────────────────────────────────────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"\s+or\s+"([^"]+)"\s+exist', c)
    if m:
        return (comps.get(m.group(1).lower(), {}).get('exists', False) or
                comps.get(m.group(2).lower(), {}).get('exists', False))

    # ── component is certified ────────────────────────────────────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"\s+is\s+certified', c)
    if m:
        return comps.get(m.group(1).lower(), {}).get('certified', False)

    # ── component do not exist / does not exist ───────────────────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"\s+do\s*(?:es)?\s*not\s+exist', c)
    if m:
        return not comps.get(m.group(1).lower(), {}).get('exists', False)

    # ── component is not certified or does not exist ──────────────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"\s+is\s+not\s+certified\s+or', c)
    if m:
        name = m.group(1).lower()
        comp = comps.get(name, {})
        return (not comp.get('certified', False)) or (not comp.get('exists', False))

    # ── EUT is class II ───────────────────────────────────────────────────────
    if 'eut is class ii' in c or 'class ii' in c:
        return 'II' in eut.get('Class', '')

    # ── EUT has "Thermoplastic case sealed" ───────────────────────────────────
    if 'thermoplastic case sealed' in c:
        case = eut.get('Case', '').lower()
        return 'thermoplastic case sealed' in case and \
               'not sealed' not in case

    # ── EUT is a "power supply" ───────────────────────────────────────────────
    if 'power supply' in c:
        return 'power supply' in eut.get('EUT', '').lower()

    # ── EUT has "marca" ───────────────────────────────────────────────────────
    if '"marca"' in c:
        return bool(data.get('Certificadora', {}).get('Registrada', ''))

    # ── EUT has "modelo" ─────────────────────────────────────────────────────
    if '"modelo"' in c:
        return bool(data.get('Certificadora', {}).get('Modelo y/o referencia tipo ', ''))

    # ── EUT has voltage / frequency / consumption mark ────────────────────────
    if 'voltage mark' in c:
        return bool(re.search(r'\d[\d\-–]+\s*V', char, re.I))
    if 'frecuency mark' in c or 'frequency mark' in c:
        return bool(re.search(r'\d[\d\-–]+\s*Hz', char, re.I))
    if 'consumption mark' in c:
        return bool(re.search(r'\d[\d,\.]+\s*[mM]?[AW]', char))

    # ── component "capacitor y" (short form without "exist") ─────────────────
    m = re.search(r'component[s]?\s+"([^"]+)"$', c)
    if m:
        return comps.get(m.group(1).lower(), {}).get('exists', False)

    # Known patterns that were already handled above
    if '6.4.8.2.2' in c:
        return False   # handled in main loop
    # Unknown condition — return False and warn
    print(f"  ⚠ Unknown condition: '{condition}'")
    return False


# ═══════════════════════════════════════════════════════════════════════════════
# Action resolver
# ═══════════════════════════════════════════════════════════════════════════════

def resolve_action(raw: str | None, data: dict, char: str) -> dict | None:
    """
    Convert an action string from the xlsx into an operation dict:
        {'op': 'set',    'text': '...'}
        {'op': 'clear'}
        {'op': 'append', 'text': '...'}
        None  → do nothing
    """
    if raw is None:
        return None

    r = raw.strip()

    # nothing / do nothing / (merged)
    if r.lower() in ('nothing.', 'nothing', 'do nothing', '(merged)', ''):
        return None

    # clear only (various phrasings)
    if re.match(r'^clear\s*\.?\s*$', r, re.I):
        return {'op': 'clear'}
    if re.match(r'^clear\s+(?:the\s+)?(?:first\s+)?cell\s*\.?\s*$', r, re.I):
        return {'op': 'clear'}

    # clear then place
    m = re.match(r'clear\s+(?:the\s+)?(?:first\s+)?cell\s+and\s+(?:the\s+)?(?:then\s+)?place\s*"([^"]*)"', r, re.I)
    if m:
        text = _resolve_text(m.group(1), data, char)
        return {'op': 'clear_set', 'text': text}

    # leaving previous, add
    m = re.match(r'leaving\s+the\s+previous\s+[^\,]+,\s*add\s+"([^"]*)"', r, re.I)
    if m:
        text = _resolve_text(m.group(1), data, char)
        return {'op': 'append', 'text': text}

    # place "text" — handle xlsx quirk where consumption text starts with 'place "EUT has'
    if re.match(r'place\s*"EUT has\s*"consumption mark', r, re.I):
        m2 = re.search(r'(\d[\d,\.]+\s*[mM]?[AW])', char)
        text = m2.group(1).strip() if m2 else 'no posee marcado de consumo'
        return {'op': 'set', 'text': text}
    m = re.match(r'place\s*"([^"]*)"', r, re.I)
    if m:
        text = _resolve_text(m.group(1), data, char)
        return {'op': 'set', 'text': text}

    print(f"  ⚠ Unknown action: '{r}'")
    return None


def _resolve_text(text: str, data: dict, char: str) -> str:
    """Replace placeholder tokens with real values from extracted data."""
    cert = data.get('Certificadora', {})

    # marca / modelo
    if text.lower() == 'marca':
        return cert.get('Registrada', '')
    if text.lower() == 'modelo':
        return cert.get('Modelo y/o referencia tipo ', '')

    # voltage mark
    if 'voltage mark' in text.lower():
        m = re.search(r'(\d[\d\-–]+\s*V[~ac]*)', char, re.I)
        return m.group(1).strip() if m else 'no posee marcado de tension'

    # frequency mark
    if 'frecuency mark' in text.lower() or 'frequency mark' in text.lower():
        m = re.search(r'(\d[\d\-–]+\s*Hz)', char, re.I)
        return m.group(1).strip() if m else 'no posee marcado de frecuencia'

    # consumption mark — also triggered when the action string itself contains 'consumption mark'
    if 'consumption mark' in text.lower():
        m = re.search(r'(\d[\d,\.]+\s*[mM]?[AW])', char)
        return m.group(1).strip() if m else 'no posee marcado de consumo'
    # Also: place "EUT has "consumption mark"..." — strip the wrapper
    if text.lower().startswith('eut has') and 'consumption mark' in text.lower():
        m = re.search(r'(\d[\d,\.]+\s*[mM]?[AW])', char)
        return m.group(1).strip() if m else 'no posee marcado de consumo'

    return text


# ═══════════════════════════════════════════════════════════════════════════════
# Row offset parser
# ═══════════════════════════════════════════════════════════════════════════════

ORDINALS = {
    'first': 1, 'second': 2, 'third': 3, 'fourth': 4,
    '4°': 4, '5°': 5, '6°': 6, '7°': 7,
    '4': 4, '5': 5, '6': 6, '7': 7,
}

def parse_clause_cell(cell: str | None) -> tuple:
    """
    Parse the 'Clause' column cell.
    Returns (clause_id, offset, is_again)

    Examples:
        '5.4.4.4'                    → ('5.4.4.4',  0, False)
        '(first row below clause)'   → (None,        1, False)
        '(first row below clause) again' → (None,    1, True)
        '(4° row below clause)'      → (None,        4, False)
    """
    if cell is None:
        return (None, 0, False)

    s = cell.strip()

    # Check for "again"
    is_again = s.lower().endswith('again')
    s_clean  = re.sub(r'\s+again$', '', s, flags=re.I).strip()

    # Nth row below clause
    m = re.match(r'\((\w+°?)\s+row\s+below\s+clause\)', s_clean, re.I)
    if m:
        word = m.group(1).lower()
        offset = ORDINALS.get(word, 0)
        return (None, offset, is_again)

    # Plain clause id
    return (s_clean, 0, is_again)


# ═══════════════════════════════════════════════════════════════════════════════
# Main translator
# ═══════════════════════════════════════════════════════════════════════════════

def translate(xlsx_path: str | Path,
              data_path:  str | Path) -> dict:
    """
    Read the xlsx and json, evaluate all conditions, and return a
    changes dict ready for modify_docx_table.py:

        {
          row_index (int): {
            "obs":        "text to set"   | None,
            "vrd":        "P" | "N"       | None,
            "clear_obs":  True            | (absent),
            "append_obs": "text to add"   | (absent),
          },
          ...
        }
    """
    with open(data_path, encoding='utf-8') as f:
        data = json.load(f)

    char = data.get('Certificadora', {}).get('Característica(s)', '')

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb['Clauses']

    # Read all rows (skip header)
    rows = [tuple(cell for cell in row)
            for row in ws.iter_rows(min_row=2, values_only=True)
            if any(cell is not None for cell in row)]

    changes   = {}   # row_idx → edit dict
    current_clause = None   # last seen real clause id
    f334_count = 0          # track F.3.3.4 row 1 vs 2
    clause_6482_result = None  # track 6.4.8.2.2 veredicto result

    def apply_action(row_idx: int, action: dict | None, field: str):
        """Write an action into changes[row_idx][field]."""
        if action is None:
            return
        entry = changes.setdefault(row_idx, {})
        if field == 'obs':
            if action['op'] == 'clear':
                entry['clear_obs'] = True
            elif action['op'] == 'clear_set':
                entry['clear_obs'] = True
                entry['obs'] = action['text']
            elif action['op'] == 'set':
                entry['obs'] = action['text']
            elif action['op'] == 'append':
                entry['append_obs'] = action.get('text', '')
        elif field == 'vrd':
            if action['op'] == 'set':
                entry['vrd'] = action['text']

    for raw_row in rows:
        clause_cell = raw_row[0]
        cond_cell   = raw_row[1] if len(raw_row) > 1 else None
        do_obs      = raw_row[2] if len(raw_row) > 2 else None
        else_obs    = raw_row[3] if len(raw_row) > 3 else None
        do_vrd      = raw_row[4] if len(raw_row) > 4 else None
        else_vrd    = raw_row[5] if len(raw_row) > 5 else None

        clause_id, offset, is_again = parse_clause_cell(clause_cell)

        # Update current clause tracker
        if clause_id is not None:
            # Special handling: F.3.3.4 appears twice (voltage + frequency)
            if clause_id == 'F.3.3.4':
                f334_count += 1
                effective_id = 'F.3.3.4' if f334_count == 1 else 'F.3.3.4b'
            else:
                f334_count   = 0
                effective_id = clause_id
            current_clause = effective_id

        # Resolve row index
        if clause_id is not None:
            base_row = CLAUSE_ROW.get(effective_id)
            if base_row is None:
                print(f"  ⚠ No row mapping for clause '{effective_id}' — skipping")
                continue
            row_idx = base_row
        else:
            # offset row below current clause
            if current_clause is None:
                print(f"  ⚠ Offset row but no current clause — skipping")
                continue
            base_row = CLAUSE_ROW.get(current_clause)
            if base_row is None:
                print(f"  ⚠ No row mapping for clause '{current_clause}' — skipping")
                continue
            row_idx = base_row + offset

        # Evaluate condition
        result = evaluate(cond_cell, data, char)

        # Special: track 6.4.8.2.2 veredicto for S.1
        if (clause_id == '6.4.8.2.2') or \
           (cond_cell and '6.4.8.2.2' in str(cond_cell)):
            # S.1 condition references 6.4.8.2.2 result
            if clause_id == '6.4.8.2.2':
                # Store what vrd will be
                clause_6482_result = 'P' if result else 'N'
            if cond_cell and '6.4.8.2.2' in str(cond_cell):
                # S.1: result depends on 6.4.8.2.2 outcome
                result = (clause_6482_result == 'P')
                # Mark as known condition so no warning is printed

        # Select obs and vrd actions
        obs_action = resolve_action(do_obs   if result else else_obs, data, char)
        vrd_action = resolve_action(do_vrd   if result else else_vrd, data, char)

        # Apply
        apply_action(row_idx, obs_action, 'obs')
        apply_action(row_idx, vrd_action, 'vrd')

        clause_label = f"{current_clause}+{offset}" if offset else current_clause
        if is_again: clause_label += ' (again)'
        cond_short = str(cond_cell)[:50] if cond_cell else '-'
        obs_short  = str(obs_action)[:40] if obs_action else '-'
        vrd_short  = str(vrd_action)[:20] if vrd_action else '-'
        verdict    = '✔ IF' if result else '✘ ELSE'
        print(f"  {verdict}  [{clause_label:<22}] row={row_idx:03d}  "
              f"obs={obs_short}  vrd={vrd_short}")

    return changes


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder containing the clauses .xlsx and extracted_data.json
    # input_folder  = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where changes.json will be written
    # output_folder = Path(r'C:\Users\yourname\Documents\job_folder')
    # ── End of configuration ──────────────────────────────────────────────────

    output_folder.mkdir(parents=True, exist_ok=True)

    # Find clauses xlsx
    xlsx_candidates = [f for f in input_folder.iterdir()
                       if f.is_file() and f.suffix.lower() == '.xlsx'
                       and 'clause' in f.name.lower()]
    if not xlsx_candidates:
        print(f"ERROR: No *clause*.xlsx found in: {input_folder}"); sys.exit(1)
    xlsx_file = xlsx_candidates[0]

    # Find extracted_data.json
    json_file = input_folder / 'extracted_data.json'
    if not json_file.exists():
        print(f"ERROR: extracted_data.json not found in: {input_folder}"); sys.exit(1)

    print(f"\n  Logic : {xlsx_file.name}")
    print(f"  Data  : {json_file.name}\n")

    changes = translate(xlsx_file, json_file)

    # Serialise — convert int keys to strings for JSON
    out = {str(k): v for k, v in sorted(changes.items())}
    out_file = output_folder / 'changes.json'
    with open(out_file, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"\n  ✔ {len(out)} row(s) → {out_file}\n")
