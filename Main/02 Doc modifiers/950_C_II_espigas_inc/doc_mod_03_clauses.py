"""
apply_clauses.py
=================
Applies the rules described in the "Kernel" / "Clauses" sheets of
950_C_II_espigas_inc_clauses_prompt.xlsx directly to word/document.xml
inside the .docx (unpack -> edit XML with lxml -> repack), the same way
doc_mod_03_clauses.py does it.

Fills "Resultados - Observaciones" and "Veredicto" cells with a single
yellow-highlighted run. Cells that get cleared also collapse their row
height to auto (minimum), per the Kernel sheet's instructions.

Clauses handled:
    1.5.6     Capacitores puenteando aislaciones (Capacitor X / Capacitor Y)
    1.5.7     Resistores puenteando aislaciones (Resistencia)
    1.5.7.1   Puenteando aislación funcional, básica o suplementaria (verdict only)
    1.5.9     Supresores (Varistor)
    1.7.1.2   Marcado de identificación (merged header, untouched)
                -> next row: marca
                -> next row: modelo
                -> next row: símbolo de clase II
    1.7.6     Identificación de los fusibles (Fusible)

Usage:
    python apply_clauses.py INPUT.docx OUTPUT.docx \
        --components component_status.json \
        --eut eut_basic.json \
        --certificadora certificadora.json

Import as module:
    from apply_clauses import apply_clauses
    apply_clauses("report.docx", eut_data, cert_data, components, "report_out.docx")
"""

import sys
import json
import shutil
import zipfile
import tempfile
import argparse
from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("ERROR: lxml not installed. Run:  pip install lxml"); sys.exit(1)


# ── Paths (easy to reach / override) ────────────────────────────────────────
# Mirrors doc_mod_03_clauses.py: pull folders/filenames from doc_mod_01_creation
# when available, otherwise fall back to sensible local defaults so this file
# can still be run/imported standalone.
try:
    import doc_mod_01_creation

    docx_folder   = doc_mod_01_creation.output_folder
    json_folder   = doc_mod_01_creation.json_folder
    output_folder = doc_mod_01_creation.output_folder

    EUT_basic_json      = doc_mod_01_creation.EUT_basic_json
    certificadora_json  = doc_mod_01_creation.certificadora_json
    componente_json     = doc_mod_01_creation.componente_existencia_json
except ImportError:
    docx_folder   = Path('.')
    json_folder   = Path('.')
    output_folder = Path('.')

    EUT_basic_json      = 'eut_basic.json'
    certificadora_json  = 'certificadora.json'
    componente_json     = 'component_status.json'


# ── Namespace ────────────────────────────────────────────────────────────────
W       = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPC = '{http://www.w3.org/XML/1998/namespace}space'


# ══════════════════════════════════════════════════════════════════════════
# XML helpers
# ══════════════════════════════════════════════════════════════════════════

def make_yellow_run(text: str) -> etree.Element:
    r = etree.fromstring(
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:rPr>'
        '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:color w:val="000000"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/>'
        '<w:highlight w:val="yellow"/>'
        '</w:rPr><w:t/></w:r>')
    t = r.find(f'{W}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set(XML_SPC, 'preserve')
    return r


def cell_text(tc: etree.Element) -> str:
    return ''.join(t.text or '' for t in tc.iter(f'{W}t')).strip()


def set_cell(tc: etree.Element, text: str):
    """Collapse the cell to a single paragraph and insert one
    yellow-highlighted run. Removes any extra paragraphs and any
    leftover content (runs, hyperlinks, etc.) so nothing from the
    original text survives."""
    paras = tc.findall(f'{W}p')
    if not paras:
        p = etree.SubElement(tc, f'{W}p')
    else:
        p = paras[0]
        for extra in paras[1:]:
            tc.remove(extra)
        for child in list(p):
            if child.tag != f'{W}pPr':
                p.remove(child)
    p.append(make_yellow_run(text))


def append_cell_line(tc: etree.Element, text: str):
    """Add a new paragraph (new line) with a yellow-highlighted run,
    without touching any existing paragraphs/runs in the cell."""
    p = etree.SubElement(tc, f'{W}p')
    p.append(make_yellow_run(text))


def clear_cell(tc: etree.Element):
    """Remove all content (runs, hyperlinks, extra paragraphs) and
    collapse row height to auto."""
    paras = tc.findall(f'{W}p')
    if paras:
        p = paras[0]
        for extra in paras[1:]:
            tc.remove(extra)
        for child in list(p):
            if child.tag != f'{W}pPr':
                p.remove(child)
    tr = tc.getparent()
    if tr is not None and tr.tag == f'{W}tr':
        trPr = tr.find(f'{W}trPr')
        if trPr is None:
            trPr = etree.SubElement(tr, f'{W}trPr')
            tr.insert(0, trPr)
        trH = trPr.find(f'{W}trHeight')
        if trH is not None:
            trPr.remove(trH)
        trH = etree.SubElement(trPr, f'{W}trHeight')
        trH.set(f'{W}val', '0')
        trH.set(f'{W}hRule', 'auto')


# ══════════════════════════════════════════════════════════════════════════
# Clause table helpers
# ══════════════════════════════════════════════════════════════════════════

def find_clause_tr(root_el: etree.Element, clause: str):
    """
    Find the <w:tr> whose first cell exactly matches clause.
    Returns (tr, parent_table_rows_list) or (None, []).
    Skips 2-cell rows (the "Listado de observaciones" table). Report rows
    have either 4 cells, or 3 when "Requisitos" and "Resultados" are
    merged (see Kernel sheet note on "(merged)" rows).
    """
    for tr in root_el.iter(f'{W}tr'):
        cells = tr.findall(f'{W}tc')
        if len(cells) < 3:
            continue
        if cell_text(cells[0]) == clause:
            for tbl in root_el.iter(f'{W}tbl'):
                tbl_trs = list(tbl.iter(f'{W}tr'))
                if tr in tbl_trs:
                    return tr, tbl_trs
    return None, []


def next_row(tbl_trs: list, current_tr) -> etree.Element | None:
    """Return the row immediately after current_tr in the table."""
    try:
        idx = tbl_trs.index(current_tr)
        if idx + 1 < len(tbl_trs):
            return tbl_trs[idx + 1]
    except ValueError:
        pass
    return None


def obs(tr: etree.Element) -> etree.Element | None:
    """Return the Resultados-Observaciones cell (index 2) of a row."""
    cells = tr.findall(f'{W}tc')
    return cells[2] if len(cells) > 2 else None


def vrd(tr: etree.Element) -> etree.Element | None:
    """Return the Veredicto cell (index 3)."""
    cells = tr.findall(f'{W}tc')
    return cells[3] if len(cells) > 3 else None


def ex(components: dict, name: str) -> bool:
    return bool(components.get(name.lower(), {}).get('exists', False))


def cert(components: dict, name: str) -> bool:
    return bool(components.get(name.lower(), {}).get('certified', False))


# ══════════════════════════════════════════════════════════════════════════
# Clause applications
# ══════════════════════════════════════════════════════════════════════════

def apply_clauses(docx_path:    str | Path,
                   eut_data:    dict,
                   cert_data:   dict,
                   components:  dict,
                   output_path: str | Path | None = None) -> Path:
    """
    Apply all clause logic to docx_path and write to output_path.

    Parameters
    ----------
    docx_path   : input .docx
    eut_data    : dict from Path_eut_basic.json
                  e.g. keys: "Norma", "Clase de aislación", "Conexión de
                  alimentación", "Carcasa del EUT", "EUT"
    cert_data   : dict from Path_certificadora_pdf.json
                  keys used here: "Marca", "Modelo"
    components  : dict from Path_exist_cert_of_component.json
                  e.g. { "capacitor x": {"exists": True, "certified": True}, ... }
    output_path : output .docx (defaults to overwriting docx_path)
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path
    output_path = Path(output_path)

    is_class2 = eut_data.get('Clase de aislación', '') == 'Clase II'

    # ── Unpack ───────────────────────────────────────────────────────────
    tmp = Path(tempfile.mkdtemp()) / 'unpacked'
    tmp.mkdir(parents=True)
    with zipfile.ZipFile(docx_path) as zf:
        zf.extractall(tmp)

    doc_path = tmp / 'word' / 'document.xml'
    doc_tree = etree.parse(str(doc_path))
    root_el  = doc_tree.getroot()

    _mod_rows: list = []
    _mod_ids:  set  = set()

    def _track(row):
        if row is None: return
        rid = id(row)
        if rid not in _mod_ids:
            _mod_ids.add(rid)
            _mod_rows.append(row)

    # ─────────────────────────────────────────────────────────────────────
    # 1.5.6  Capacitores puenteando aislaciones (Capacitor X / Capacitor Y)
    # ─────────────────────────────────────────────────────────────────────
    tr, _ = find_clause_tr(root_el, '1.5.6')
    if tr is not None:
        cap_x = ex(components, 'capacitor x')
        cap_y = ex(components, 'capacitor y')
        lines = []
        if cap_x: lines.append('Capacitor X certificado entre polos de alimentación con aislación X2')
        if cap_y: lines.append('Capacitor Y certificado entre primario y secundario con aislación Y1.')
        c = obs(tr)
        if c is not None and lines:
            set_cell(c, lines[0])
            for extra in lines[1:]:
                append_cell_line(c, extra)
        # else: "nothing" - leave Resultados untouched
        c = vrd(tr)
        if c is not None: set_cell(c, 'P' if cap_y else 'N')  # last evaluated rule wins
        _track(tr)
        print(f"  \u2714 1.5.6: {'P' if cap_y else 'N'}")

    # ─────────────────────────────────────────────────────────────────────
    # 1.5.7  Resistores puenteando aislaciones (Resistencia)
    # ─────────────────────────────────────────────────────────────────────
    tr, _ = find_clause_tr(root_el, '1.5.7')
    if tr is not None:
        resistencia = ex(components, 'resistencia')
        c = obs(tr)
        if c is not None:
            set_cell(c, 'Resistencia entre polos de alimentación después del fusible.' if resistencia else '---')
        c = vrd(tr)
        if c is not None: set_cell(c, 'P' if resistencia else 'N')
        _track(tr)
        print(f"  \u2714 1.5.7: {'P' if resistencia else 'N'}")

    # ─────────────────────────────────────────────────────────────────────
    # 1.5.7.1  Puenteando aislación funcional, básica o suplementaria
    # (verdict only - Resultados untouched either way)
    # ─────────────────────────────────────────────────────────────────────
    tr, _ = find_clause_tr(root_el, '1.5.7.1')
    if tr is not None:
        resistencia = ex(components, 'resistencia')
        c = vrd(tr)
        if c is not None: set_cell(c, 'P' if resistencia else 'N')
        _track(tr)
        print(f"  \u2714 1.5.7.1: {'P' if resistencia else 'N'}")

    # ─────────────────────────────────────────────────────────────────────
    # 1.5.9  Supresores (Varistor)
    # ─────────────────────────────────────────────────────────────────────
    tr, _ = find_clause_tr(root_el, '1.5.9')
    if tr is not None:
        varistor = ex(components, 'varistor')
        c = obs(tr)
        if c is not None and varistor:
            set_cell(c, 'Varistor certificado. (ver tabla 1.5.1)')
        # else: "nothing" - leave Resultados untouched
        c = vrd(tr)
        if c is not None: set_cell(c, 'P' if varistor else 'N')
        _track(tr)
        print(f"  \u2714 1.5.9: {'P' if varistor else 'N'}")

    # ─────────────────────────────────────────────────────────────────────
    # 1.7.1.2  Marcado de identificación (merged header row - untouched)
    #   row+1: marca
    #   row+2: modelo
    #   row+3: símbolo de clase II
    # ─────────────────────────────────────────────────────────────────────
    tr0, tbl_trs = find_clause_tr(root_el, '1.7.1.2')
    if tr0 is not None:
        marca  = (cert_data or {}).get('Marca', '') or ''
        modelo = (cert_data or {}).get('Modelo', '') or ''

        tr1 = next_row(tbl_trs, tr0)
        if tr1 is not None:
            c = obs(tr1)
            if c is not None and marca: set_cell(c, marca)
            c = vrd(tr1)
            if c is not None: set_cell(c, 'P' if marca else 'F')
            _track(tr1)

        tr2 = next_row(tbl_trs, tr1) if tr1 is not None else None
        if tr2 is not None:
            c = obs(tr2)
            if c is not None and modelo: set_cell(c, modelo)
            c = vrd(tr2)
            if c is not None: set_cell(c, 'P' if modelo else 'F')
            _track(tr2)

        tr3 = next_row(tbl_trs, tr2) if tr2 is not None else None
        if tr3 is not None:
            c = obs(tr3)
            if c is not None:
                if is_class2: set_cell(c, 'Posee símbolo clase II ')
                else: clear_cell(c)
            c = vrd(tr3)
            if c is not None: set_cell(c, 'P' if is_class2 else 'N')
            _track(tr3)

        print(f"  \u2714 1.7.1.2: marca='{marca}' modelo='{modelo}' clase II={'P' if is_class2 else 'N'}")

    # ─────────────────────────────────────────────────────────────────────
    # 1.7.6  Identificación de los fusibles (Fusible)
    # ─────────────────────────────────────────────────────────────────────
    tr, _ = find_clause_tr(root_el, '1.7.6')
    if tr is not None:
        fusible = ex(components, 'fusible')
        c = obs(tr)
        if c is not None and not fusible:
            clear_cell(c)
        # else: "nothing" - leave Resultados untouched
        c = vrd(tr)
        if c is not None: set_cell(c, 'P' if fusible else 'N')
        _track(tr)
        print(f"  \u2714 1.7.6: {'P' if fusible else 'N'}")

    # ── Save output ─────────────────────────────────────────────────────
    doc_tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)
    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(tmp.rglob('*')):
            if f.is_file():
                zf.write(f, f.relative_to(tmp))
    print(f"\n  \u2714 Output: {output_path}")

    shutil.rmtree(tmp)
    return output_path


# ══════════════════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('input_docx', nargs='?',
                         help='Defaults to looking up the filename via '
                              'doc_mod_numero_informe.json in json_folder')
    parser.add_argument('output_docx', nargs='?',
                         help='Defaults to <docx_folder>/<input filename>')
    parser.add_argument('--components', default=None,
                         help=f'Path_exist_cert_of_component JSON (default: {componente_json})')
    parser.add_argument('--eut', default=None,
                         help=f'Path_eut_basic JSON (default: {EUT_basic_json})')
    parser.add_argument('--certificadora', default=None,
                         help=f'Path_certificadora_pdf JSON (default: {certificadora_json})')
    args = parser.parse_args()

    components_path = Path(args.components) if args.components else Path(json_folder) / componente_json
    eut_path         = Path(args.eut) if args.eut else Path(json_folder) / EUT_basic_json
    cert_path         = Path(args.certificadora) if args.certificadora else Path(json_folder) / certificadora_json

    if args.input_docx:
        docx_file = Path(args.input_docx)
    else:
        informe_json = Path(json_folder) / "doc_mod_numero_informe.json"
        if not informe_json.exists():
            print(f"ERROR: '{informe_json}' not found. Pass an input .docx explicitly, "
                  f"or run docx_modify_creation.py first.")
            sys.exit(1)
        with open(informe_json, encoding='utf-8') as f:
            _inf = json.load(f)
        docx_filename = _inf.get("filename")
        if not docx_filename:
            print(f"ERROR: 'filename' key missing in '{informe_json}'."); sys.exit(1)
        docx_file = Path(docx_folder) / docx_filename
        if not docx_file.exists():
            print(f"ERROR: '{docx_file}' not found."); sys.exit(1)
        print(f"\n  Found docx: {docx_file.name}")

    output_file = Path(args.output_docx) if args.output_docx else Path(output_folder) / docx_file.name

    missing = [p for p in (components_path, eut_path, cert_path) if not p.exists()]
    if missing:
        print(f"ERROR: Missing JSON file(s): {', '.join(str(p) for p in missing)}")
        sys.exit(1)

    with open(components_path, encoding='utf-8') as f:
        components = json.load(f)
    with open(eut_path, encoding='utf-8') as f:
        eut_data = json.load(f)
    with open(cert_path, encoding='utf-8') as f:
        cert_data = json.load(f)

    print(f"\n  File   : {docx_file.name}\n")

    apply_clauses(docx_file, eut_data, cert_data, components, output_file)
    print("\nDone.\n")
