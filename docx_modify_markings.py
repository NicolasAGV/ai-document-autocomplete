from pathlib import Path

docx_folder = Path(r"C:\Users\nicol\OneDrive\Desktop\MASTER Prueba AI Claude IEC 62368\Prueba")
json_folder = Path(r"C:\Users\nicol\OneDrive\Desktop\MASTER Prueba AI Claude IEC 62368\Prueba\output")
output_folder = docx_folder / 'output'

"""
modify_docx_fields.py  —  Part 4
==================================
Fills all text fields in the .docx report with yellow highlight:

  Headers (all pages)
      Informe No.: → Número de Informe

  Page 1  (Table 0)
      Informe de Referencia No..:  → Número de Informe  (form field)
      Nombre del Solicitante:      → Cliente
      Dirección:                   → Dirección  (the empty one after Solicitante)
      Descripción del ítem de ensayo: → Descripción del ítem de ensayo
      Marca Registrada :           → Registrada
      Modelo y/o referencia tipo : → Modelo y/o referencia tipo
      Característica(s)…:          → Característica(s)

  Cert.Pag.  (Table 3)
      Certificadora:               → Certificadora
      Solicitud:                   → Solicitud
      Identificación:              → Identificación

  Particularidades  (Table 7)
      Conexión a la alimentación…: → Plug  (from EUT config)
      Clase de equipo .:           → Class (from EUT config)
      Fecha de recepción…:         → Fecha de Alta

Usage:
    python modify_docx_fields.py  <docx>  <sistema.json|dict>  <cert.json|dict>  <eut.json|dict>  [output.docx]

    In practice, called from run_all.py which passes dicts directly.

    CLI example (pass JSON files):
        python modify_docx_fields.py report.docx sistema.json cert.json eut.json report_fields.docx

    JSON format for each file:
        sistema.json  : {"Número de Informe": "...", "Fecha de Alta": "...",
                         "Cliente": "...", "Dirección": "..."}
        cert.json     : {"Certificadora": "...", "Solicitud": "...",
                         "Identificación": "...",
                         "Descripción del ítem de ensayo": "...",
                         "Registrada": "...",
                         "Modelo y/o referencia tipo ": "...",
                         "Característica(s)": "..."}
        eut.json      : {"Class": "Clase II", "Plug": "Fichas incorporadas",
                         "Case": "Thermoplastic case sealed", "EUT": "Power supply"}

Install dependencies (once):
    pip install lxml

Import as module:
    from modify_docx_fields import fill_fields
    fill_fields("report.docx", sis_data, cert_data, eut_data, "report_fields.docx")
"""

import sys
import json
import shutil
import zipfile
import tempfile
from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("ERROR: lxml not installed. Run:  pip install lxml"); sys.exit(1)


# ── Namespace constants ───────────────────────────────────────────────────────
W       = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPC = '{http://www.w3.org/XML/1998/namespace}space'


# ═══════════════════════════════════════════════════════════════════════════════
# XML helpers
# ═══════════════════════════════════════════════════════════════════════════════

def make_yellow_run(text: str) -> etree.Element:
    """Create a run with yellow highlight."""
    r = etree.fromstring(
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:rPr>'
        '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:color w:val="000000"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/>'
        '<w:highlight w:val="yellow"/>'
        '</w:rPr>'
        '<w:t/>'
        '</w:r>')
    t = r.find(f'{W}t')
    t.text = text
    if text and (text[0] == ' ' or text[-1] == ' '):
        t.set(XML_SPC, 'preserve')
    return r


def cell_text(tc: etree.Element) -> str:
    return ''.join(t.text or '' for t in tc.iter(f'{W}t')).strip()


def set_cell(tc: etree.Element, text: str):
    """Clear all runs in cell and insert a single yellow run."""
    for p in tc.findall(f'{W}p'):
        for r in list(p.findall(f'{W}r')):
            p.remove(r)
    paras = tc.findall(f'{W}p')
    p = paras[0] if paras else etree.SubElement(tc, f'{W}p')
    p.append(make_yellow_run(text))


def find_value_cell(root_el: etree.Element,
                    label: str,
                    exact: bool = False) -> etree.Element | None:
    """
    Find the cell immediately to the RIGHT of a label cell.
    exact=True  → cell text must equal label exactly
    exact=False → cell text must CONTAIN label
    """
    for tr in root_el.iter(f'{W}tr'):
        cells = tr.findall(f'{W}tc')
        for i, tc in enumerate(cells):
            txt = cell_text(tc)
            hit = (txt == label) if exact else (label in txt)
            if hit and i + 1 < len(cells):
                return cells[i + 1]
    return None


# ═══════════════════════════════════════════════════════════════════════════════
# 1 — Headers: replace "nro informe" with the actual informe number
# ═══════════════════════════════════════════════════════════════════════════════

def _update_headers(word_dir: Path, informe: str):
    """
    Replace 'nro informe' placeholder in all header XML files.
    The text is split across multiple runs, so we collect the run sequence
    that together spell 'nro informe' and replace them with one yellow run.
    """
    for hfile in sorted(word_dir.glob('header*.xml')):
        htree = etree.parse(str(hfile))
        hroot = htree.getroot()
        changed = False

        for p in hroot.iter(f'{W}p'):
            full = ''.join(t.text or '' for t in p.iter(f'{W}t'))
            if 'nro informe' not in full.lower():
                continue

            # Collect the run sequence that forms "nro informe"
            runs = p.findall(f'{W}r')
            seq = []
            collecting = False
            for run in runs:
                rt = ''.join(t.text or '' for t in run.findall(f'{W}t'))
                if 'nro' in rt.lower():
                    collecting = True
                if collecting:
                    seq.append(run)
                if collecting and 'informe' in rt.lower():
                    break

            if seq:
                # Remove collected runs and insert one yellow run
                for run in seq:
                    p.remove(run)
                p.append(make_yellow_run(informe))
                changed = True

        if changed:
            htree.write(str(hfile),
                        xml_declaration=True, encoding='UTF-8', standalone=True)
            print(f"  ✔ Header updated: {hfile.name}")


# ═══════════════════════════════════════════════════════════════════════════════
# 2 — Body form field: "nro informe" inside a FORMTEXT field
# ═══════════════════════════════════════════════════════════════════════════════

def _update_body_field(root_el: etree.Element, informe: str):
    """
    The 'Informe de Referencia No..' cell contains a FORMTEXT field.
    Between the fldChar separate and end tags the text run holds 'nro informe'.
    Replace it with the actual informe number and add yellow highlight.
    Also update the w:default value of the field.
    """
    for p in root_el.iter(f'{W}p'):
        in_field = False
        for r in p.findall(f'{W}r'):
            fc = r.find(f'{W}fldChar')
            if fc is not None and fc.get(f'{W}fldCharType') == 'separate':
                in_field = True
                continue
            if fc is not None and fc.get(f'{W}fldCharType') == 'end':
                in_field = False
            if in_field:
                for t in r.findall(f'{W}t'):
                    if t.text and 'nro informe' in t.text.lower():
                        t.text = informe
                        # Add yellow highlight to this run
                        rpr = r.find(f'{W}rPr')
                        if rpr is None:
                            rpr = etree.SubElement(r, f'{W}rPr')
                            r.insert(0, rpr)
                        hl = rpr.find(f'{W}highlight')
                        if hl is None:
                            hl = etree.SubElement(rpr, f'{W}highlight')
                        hl.set(f'{W}val', 'yellow')
                        print(f"  ✔ Body form field: nro informe → {informe}")

    # Also update the <w:default> value inside the field definition
    for d in root_el.iter(f'{W}default'):
        if d.get(f'{W}val') == 'nro informe':
            d.set(f'{W}val', informe)


# ═══════════════════════════════════════════════════════════════════════════════
# 3 — Page 1: fill all simple label → value cells
# ═══════════════════════════════════════════════════════════════════════════════

def _update_page1(root_el: etree.Element,
                  sis_data: dict, cert_data: dict):
    """
    Fill the Page 1 fields.
    Each field is identified by a partial label string match.
    """
    # Map: label substring → value
    # Using partial match (exact=False) so minor whitespace differences don't matter
    fields = {
        'Nombre del Solicitante':         sis_data.get('Cliente', ''),
        'Descripción del ítem de ensayo': cert_data.get('Descripción del ítem de ensayo', ''),
        'Marca Registrada':               cert_data.get('Registrada', ''),
        'Modelo y/o referencia tipo':     cert_data.get('Modelo y/o referencia tipo ', ''),
        'Característica(s)':              cert_data.get('Característica(s)', ''),
    }

    for label, value in fields.items():
        tc = find_value_cell(root_el, label)
        if tc is not None:
            set_cell(tc, value)
            print(f"  ✔ {label}")
        else:
            print(f"  ⚠ NOT FOUND: {label}")

    # Dirección: the first empty value cell after "Nombre del Solicitante"
    # (there are two "Dirección" rows; the one after Solicitante is the client's)
    direccion   = sis_data.get('Dirección', '')
    found_solicitante = False
    for tr in root_el.iter(f'{W}tr'):
        cells = tr.findall(f'{W}tc')
        for i, tc in enumerate(cells):
            txt = cell_text(tc)
            if 'Nombre del Solicitante' in txt:
                found_solicitante = True
            if found_solicitante and 'Dirección' in txt and i + 1 < len(cells):
                target = cells[i + 1]
                if not cell_text(target):          # only fill if currently empty
                    set_cell(target, direccion)
                    print(f"  ✔ Dirección (after Solicitante)")
                    found_solicitante = False
                    break


# ═══════════════════════════════════════════════════════════════════════════════
# 4 — Cert.Pag.: fill Certificadora, Solicitud, Identificación
# ═══════════════════════════════════════════════════════════════════════════════

def _update_cert_pag(root_el: etree.Element, cert_data: dict):
    """
    Fill the Cert.Pag. table fields.
    Labels end with ':' so we use exact matching to avoid hitting
    other rows that contain these words.
    """
    fields = {
        'Certificadora:': cert_data.get('Certificadora', ''),
        'Solicitud:':     cert_data.get('Solicitud', ''),
        'Identificación:':cert_data.get('Identificación', ''),
    }
    for label, value in fields.items():
        tc = find_value_cell(root_el, label, exact=True)
        if tc is not None:
            set_cell(tc, value)
            print(f"  ✔ {label}")
        else:
            print(f"  ⚠ NOT FOUND: {label}")


# ═══════════════════════════════════════════════════════════════════════════════
# 5 — Particularidades: Conexión, Clase de equipo, Fecha de recepción
# ═══════════════════════════════════════════════════════════════════════════════

def _update_particularidades(root_el: etree.Element,
                              sis_data: dict, eut_data: dict):
    """
    Fill the Particularidades del ítem de ensayo table.
    Labels contain long strings with dots/spaces; partial match is fine.
    """
    fields = {
        'Conexión a la alimentación': eut_data.get('Plug', ''),
        'Clase de equipo':            eut_data.get('Class', ''),
        'Fecha de recepción del ítem':sis_data.get('Fecha de Alta', ''),
    }
    for label, value in fields.items():
        tc = find_value_cell(root_el, label)
        if tc is not None:
            set_cell(tc, value)
            print(f"  ✔ {label}")
        else:
            print(f"  ⚠ NOT FOUND: {label}")


# ═══════════════════════════════════════════════════════════════════════════════
# Docx unpack / pack helpers
# ═══════════════════════════════════════════════════════════════════════════════

def _unpack(docx: Path, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx) as zf:
        zf.extractall(dest)


def _pack(unpacked: Path, out_docx: Path):
    if out_docx.exists():
        out_docx.unlink()
    with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(unpacked.rglob('*')):
            if f.is_file():
                zf.write(f, f.relative_to(unpacked))


# ═══════════════════════════════════════════════════════════════════════════════
# Public API
# ═══════════════════════════════════════════════════════════════════════════════

def fill_fields(docx_path:   str | Path,
                sis_data:    dict,
                cert_data:   dict,
                eut_data:    dict,
                output_path: str | Path | None = None) -> Path:
    """
    Fill all text fields in docx_path and write to output_path.

    Parameters
    ----------
    docx_path   : path to the input .docx
    sis_data    : dict from extract_sistema.extract()
                  keys: Número de Informe, Fecha de Alta, Cliente, Dirección
    cert_data   : dict from extract_certificadora.extract()
                  keys: Certificadora, Solicitud, Identificación,
                        Descripción del ítem de ensayo, Registrada,
                        Modelo y/o referencia tipo , Característica(s)
    eut_data    : dict with EUT info
                  keys: Class, Plug, Case, EUT
    output_path : path to write the modified .docx
                  defaults to <docx_stem>_fields.docx

    Returns
    -------
    Path to the output .docx
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / (docx_path.stem + '_fields.docx')
    output_path = Path(output_path)

    informe = sis_data.get('Número de Informe', '')

    # ── Unpack ────────────────────────────────────────────────────────────────
    tmp = Path(tempfile.mkdtemp()) / 'unpacked'
    _unpack(docx_path, tmp)

    word_dir  = tmp / 'word'
    doc_path  = word_dir / 'document.xml'

    doc_tree = etree.parse(str(doc_path))
    root_el  = doc_tree.getroot()

    # ── Apply changes ─────────────────────────────────────────────────────────
    print("\n  — Headers —")
    _update_headers(word_dir, informe)

    print("\n  — Body form field —")
    _update_body_field(root_el, informe)

    print("\n  — Page 1 —")
    _update_page1(root_el, sis_data, cert_data)

    print("\n  — Cert.Pag. —")
    _update_cert_pag(root_el, cert_data)

    print("\n  — Particularidades —")
    _update_particularidades(root_el, sis_data, eut_data)

    # ── Save document.xml ─────────────────────────────────────────────────────
    doc_tree.write(str(doc_path),
                   xml_declaration=True, encoding='UTF-8', standalone=True)

    # ── Pack ──────────────────────────────────────────────────────────────────
    _pack(tmp, output_path)
    shutil.rmtree(tmp)

    print(f"\n  ✔ Output: {output_path}")
    return output_path


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder where the input .docx is located
    # docx_folder = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where sistema.json, certificadora.json and eut.json are located
    # json_folder = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where the output .docx will be written
    # output_folder = docx_folder / 'output'
    # ── End of configuration ──────────────────────────────────────────────────

    output_folder.mkdir(parents=True, exist_ok=True)

    # Find input .docx: prefer one with "mod" in the name, fall back to any .docx
    mod_candidates = [f for f in docx_folder.iterdir()
                      if f.is_file()
                      and f.suffix.lower() == '.docx'
                      and 'mod' in f.stem.lower()]
    any_candidates = [f for f in docx_folder.iterdir()
                      if f.is_file()
                      and f.suffix.lower() == '.docx']

    if mod_candidates:
        docx_file = mod_candidates[0]
        print(f"\n  Found modified docx: {docx_file.name}")
    elif any_candidates:
        docx_file = any_candidates[0]
        print(f"\n  No 'mod' docx found — using: {docx_file.name}")
    else:
        print(f"ERROR: No .docx found in: {docx_folder}")
        sys.exit(1)

    # Output filename: ensure "mod" is present, then append "_mrk"
    #   368_C_II.docx              → 368_C_II_mod_mrk.docx
    #   368_C_II_mod_fields.docx   → 368_C_II_mod_fields_mrk.docx
    stem = docx_file.stem
    if 'mod' not in stem.lower():
        stem = stem + '_mod'
    output_file = output_folder / (stem + '_mrk.docx')

    # Load JSON files from json_folder
    missing = []
    for fname in ['sistema.json', 'certificadora.json', 'eut.json']:
        if not (json_folder / fname).exists():
            missing.append(fname)
    if missing:
        print(f"ERROR: Missing files in '{json_folder}': {', '.join(missing)}")
        sys.exit(1)

    with open(json_folder / 'sistema.json',       encoding='utf-8') as f: sis_data  = json.load(f)
    with open(json_folder / 'certificadora.json', encoding='utf-8') as f: cert_data = json.load(f)
    with open(json_folder / 'eut.json',           encoding='utf-8') as f: eut_data  = json.load(f)

    print(f"  Input   : {docx_file.name}")
    print(f"  Output  : {output_file.name}")
    print(f"  Informe : {sis_data.get('Número de Informe', '')}")
    print()

    fill_fields(docx_file, sis_data, cert_data, eut_data, output_file)
    print("\nDone.\n")
