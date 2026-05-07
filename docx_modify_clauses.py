from pathlib import Path

docx_folder   = Path(r"C:\Users\nicol\OneDrive\Desktop\MASTER Prueba AI Claude IEC 62368\Prueba\output")
json_folder   = Path(r"C:\Users\nicol\OneDrive\Desktop\MASTER Prueba AI Claude IEC 62368\Prueba\output")
output_folder = docx_folder 

"""
modify_docx_clauses.py  —  Part 5
===================================
Applies all clause logic to the .docx report.
Fills Resultados–Observaciones and Veredicto cells with yellow highlight.

Clauses handled:
    5.4.4.4   Optoacoplador
    5.4.4.6.2 Material de lámina delgada (transformer insulation)
    5.5.2     Capacitores y unidades RC
    5.5.2.2   Protecciones contra descarga de capacitores
    5.5.7     SPD's (Varistor)
    5.7.2.1   Medición de corriente de contacto
    6.4.8.2.2 Cubierta contra fuego
    F.3.2.1   Identificación del fabricante
    F.3.2.2   Identificación del modelo
    F.3.3.4   Tensión nominal / Frecuencia nominal  (two rows)
    F.3.3.6   Rango de corriente o potencia
    F.4       Instrucciones
    G.2       Relés  (merged header)
    G.2.1     Relés general
    G.3.4     Fusible
    G.6       Aislación de alambre  (merged header)
    G.6.1     Transformer / Aislación
    G.8       Varistores  (merged header)
    G.8.1     Varistores general
    G.11      Capacitor y unidades RC  (merged header)
    G.11.1    Capacitor X / Y
    G.12      Optoacopladores  (merged header + data row)
    P.1       Requisitos generales
    P.2.2     Protección objetos extraños  (two rows)
    S.1       Ensayo de inflamabilidad  (multi-row block)

Usage:
    python modify_docx_clauses.py

    Edit the Configuration block:
        docx_folder   — folder containing the input .docx
        json_folder   — folder containing eut.json and certificadora.json
        output_folder — folder where the output .docx is written
                        (output replaces input in docx_folder)

Install dependencies (once):
    pip install lxml

Import as module:
    from modify_docx_clauses import apply_clauses
    apply_clauses("report.docx", eut_data, cert_data, "report_cla.docx")
"""

import re
import sys
import json
import shutil
import zipfile
import tempfile

try:
    from lxml import etree
except ImportError:
    print("ERROR: lxml not installed. Run:  pip install lxml"); sys.exit(1)


# ── Namespace ─────────────────────────────────────────────────────────────────
W       = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XML_SPC = '{http://www.w3.org/XML/1998/namespace}space'


# ═══════════════════════════════════════════════════════════════════════════════
# XML helpers
# ═══════════════════════════════════════════════════════════════════════════════

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
    """Clear all runs and insert a single yellow-highlighted run."""
    for p in tc.findall(f'{W}p'):
        for r in list(p.findall(f'{W}r')):
            p.remove(r)
    paras = tc.findall(f'{W}p')
    p = paras[0] if paras else etree.SubElement(tc, f'{W}p')
    p.append(make_yellow_run(text))


def clear_cell(tc: etree.Element):
    """Remove all runs and collapse row height to auto."""
    for p in tc.findall(f'{W}p'):
        for r in list(p.findall(f'{W}r')):
            p.remove(r)
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


# ═══════════════════════════════════════════════════════════════════════════════
# Clause table helpers
# ═══════════════════════════════════════════════════════════════════════════════

def find_main_clause_table(root_el: etree.Element) -> etree.Element | None:
    """
    Find the main clause table — the one that contains S.1, P.2.2,
    G.6 and F.4 all with at least 3-cell rows.
    """
    for tbl in root_el.iter(f'{W}tbl'):
        found = set()
        for tr in tbl.iter(f'{W}tr'):
            cells = tr.findall(f'{W}tc')
            if not cells: continue
            ft = cell_text(cells[0])
            if ft in ('S.1', 'P.2.2', 'G.6', 'F.4') and len(cells) >= 3:
                found.add(ft)
        if len(found) >= 4:
            return tbl
    return None


def find_clause_tr(root_el: etree.Element, clause: str):
    """
    Find the <w:tr> whose first cell exactly matches clause.
    Uses root.iter() so nested tables are handled correctly.
    Returns (tr, parent_table_rows_list) or (None, []).
    """
    for tr in root_el.iter(f'{W}tr'):
        cells = tr.findall(f'{W}tc')
        if not cells: continue
        if cell_text(cells[0]) == clause:
            # Get all rows of the parent table
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
    """Return the Resultados–Observaciones cell (index 2) of a row."""
    cells = tr.findall(f'{W}tc')
    return cells[2] if len(cells) > 2 else None


def vrd(tr: etree.Element) -> etree.Element | None:
    """
    Return the Veredicto cell.
    Normal rows: index 3.
    Merged-header rows (3 cells, gridSpan=2 on cell 1): index 2.
    """
    cells = tr.findall(f'{W}tc')
    if len(cells) == 3:
        return cells[2]   # merged header: clause | merged(desc+obs) | veredicto
    if len(cells) >= 4:
        return cells[3]
    return None


def ex(components: dict, name: str) -> bool:
    return components.get(name.lower(), {}).get('exists', False)


def cert(components: dict, name: str) -> bool:
    return components.get(name.lower(), {}).get('certified', False)


# ═══════════════════════════════════════════════════════════════════════════════
# Clause applications
# ═══════════════════════════════════════════════════════════════════════════════

def apply_clauses(docx_path:    str | Path,
                  eut_data:     dict,
                  cert_data:    dict,
                  components:   dict,
                  output_path:  str | Path | None = None) -> Path:
    """
    Apply all clause logic to docx_path and write to output_path.

    Parameters
    ----------
    docx_path   : input .docx
    eut_data    : dict from extract_eut.py
                  keys: Class, Plug, Case, EUT
    cert_data   : dict from extract_certificadora.py
                  keys: Registrada, Modelo y/o referencia tipo , Característica(s), ...
    components  : dict  { "optoacoplador": {"exists": True, "certified": True}, ... }
    output_path : output .docx (defaults to <stem>_cla.docx)
    """
    docx_path = Path(docx_path)
    if output_path is None:
        output_path = docx_path.parent / (docx_path.stem + '_cla.docx')
    output_path = Path(output_path)

    is_class2  = 'II'      in eut_data.get('Class', '')
    is_sealed  = 'sealed'  in eut_data.get('Case',  '').lower() and \
                 'not sealed' not in eut_data.get('Case', '').lower()
    is_psu     = 'power supply' in eut_data.get('EUT', '').lower()

    char   = cert_data.get('Característica(s)', '')
    volt_m = re.search(r'(\d[\d\-–]+\s*V[~ac]*)', char, re.I)
    freq_m = re.search(r'(\d[\d\-–]+\s*Hz)',       char, re.I)
    amp_m  = re.search(r'(\d[\d,\.]+\s*[mM]?A)',   char)
    volt   = volt_m.group(1).strip() if volt_m else 'no posee marcado de tension'
    freq   = freq_m.group(1).strip() if freq_m else 'no posee marcado de frecuencia'
    amp    = amp_m.group(1).strip()  if amp_m  else 'no posee marcado de consumo'

    # ── Unpack ────────────────────────────────────────────────────────────────
    tmp = Path(tempfile.mkdtemp()) / 'unpacked'
    tmp.mkdir(parents=True)
    with zipfile.ZipFile(docx_path) as zf:
        zf.extractall(tmp)

    doc_path = tmp / 'word' / 'document.xml'
    doc_tree = etree.parse(str(doc_path))
    root_el  = doc_tree.getroot()

    # ── Locate clause table ───────────────────────────────────────────────────
    clause_tbl = find_main_clause_table(root_el)
    if clause_tbl is None:
        print("  ⚠ Main clause table not found!"); return output_path

    clause_trs = list(clause_tbl.iter(f'{W}tr'))

    def _find(clause):
        """Find clause row within the main clause table."""
        for tr in clause_trs:
            cells = tr.findall(f'{W}tc')
            if cells and cell_text(cells[0]) == clause:
                return tr
        return None

    def _next(tr):
        """Next row in the clause table."""
        try:
            i = clause_trs.index(tr)
            return clause_trs[i + 1] if i + 1 < len(clause_trs) else None
        except ValueError:
            return None

    def _find_nested(clause):
        """
        Find clause row using root.iter() — for clauses in nested tables
        (5.x, 6.x) that are inside a sub-table within the main clause table.
        Returns (tr, parent_trs_list).
        """
        return find_clause_tr(root_el, clause)

    # ─────────────────────────────────────────────────────────────────────────
    # 5.4.4.4  Optoacoplador
    # ─────────────────────────────────────────────────────────────────────────
    tr, tbl_trs = _find_nested('5.4.4.4')
    if tr is not None:
        if ex(components, 'optoacoplador'):
            c = obs(tr); 
            if c is not None: set_cell(c, 'Optoacoplador certificado. (Ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ 5.4.4.4: {'P' if ex(components,'optoacoplador') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # 5.4.4.6.2  Material de lámina delgada
    # ─────────────────────────────────────────────────────────────────────────
    tr, tbl_trs = _find_nested('5.4.4.6.2')
    if tr is not None:
        idx    = tbl_trs.index(tr)
        tr2    = tbl_trs[idx + 1] if idx + 1 < len(tbl_trs) else None
        ais_ex = ex(components, 'aislación del transformador')
        if ais_ex:
            c = obs(tr);  
            if c is not None: clear_cell(c)
            if tr2 is not None:
                c = obs(tr2); 
                if c is not None: clear_cell(c); set_cell(c, '---')
            for r in ([tr] + ([tr2] if tr2 else [])):
                c = vrd(r); 
                if c is not None: set_cell(c, 'N')
        else:
            c = obs(tr);  
            if c is not None: clear_cell(c); set_cell(c, 'Cada capa pasa el ensayo de rigidez dieléctrica a 3000Vca.')
            if tr2 is not None:
                c = obs(tr2); 
                if c is not None: clear_cell(c); set_cell(c, '2 capas entre bobinado primario y secundario')
            for r in ([tr] + ([tr2] if tr2 else [])):
                c = vrd(r); 
                if c is not None: set_cell(c, 'P')
        print(f"  ✔ 5.4.4.6.2: {'N' if ais_ex else 'P'}")

    # ─────────────────────────────────────────────────────────────────────────
    # 5.5.2  Capacitores y unidades RC
    # ─────────────────────────────────────────────────────────────────────────
    tr, _ = _find_nested('5.5.2')
    if tr is not None:
        has_cap = ex(components,'capacitor x') or ex(components,'capacitor y')
        if has_cap:
            c = obs(tr); 
            if c is not None: set_cell(c, 'Cumple Anexo G.11')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ 5.5.2: {'P' if has_cap else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # 5.5.2.2  Protecciones contra la descarga de capacitores
    # ─────────────────────────────────────────────────────────────────────────
    tr, _ = _find_nested('5.5.2.2')
    if tr is not None:
        if ex(components, 'capacitor x'):
            c = obs(tr); 
            if c is not None: set_cell(c, ' 12Vp ; 2s')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = obs(tr); 
            if c is not None: set_cell(c, '---')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ 5.5.2.2: {'P' if ex(components,'capacitor x') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # 5.5.7  SPD's (Varistor)
    # ─────────────────────────────────────────────────────────────────────────
    tr, _ = _find_nested('5.5.7')
    if tr is not None:
        if ex(components, 'varistor'):
            c = obs(tr); 
            if c is not None: set_cell(c, 'Varistor certificado.(Ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = obs(tr); 
            if c is not None: clear_cell(c)
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ 5.5.7: {'P' if ex(components,'varistor') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # 5.7.2.1  Medición de corriente de contacto
    # ─────────────────────────────────────────────────────────────────────────
    tr, _ = _find_nested('5.7.2.1')
    if tr is not None:
        txt = 'Carcasa termoplástica: U2 = 6mV  Salida: U2 = 12mV' if is_class2 \
              else 'Tierra: U2 = 62mV'
        c = obs(tr); 
        if c is not None: set_cell(c, txt)
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P')
        print(f"  ✔ 5.7.2.1: P")

    # ─────────────────────────────────────────────────────────────────────────
    # 6.4.8.2.2  Cubierta contra fuego
    # ─────────────────────────────────────────────────────────────────────────
    tr, _ = _find_nested('6.4.8.2.2')
    clause_6482_P = False
    if tr is not None:
        if is_sealed:
            c = obs(tr); 
            if c is not None: set_cell(c, 'Cumple S.1')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
            clause_6482_P = True
        else:
            c = obs(tr); 
            if c is not None: clear_cell(c)
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ 6.4.8.2.2: {'P' if is_sealed else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # F.3.2.1 / F.3.2.2  Fabricante / Modelo
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('F.3.2.1')
    if tr is not None:
        c = obs(tr); 
        if c is not None: set_cell(c, cert_data.get('Registrada', ''))
        print(f"  ✔ F.3.2.1: {cert_data.get('Registrada','')}")

    tr = _find('F.3.2.2')
    if tr is not None:
        c = obs(tr); 
        if c is not None: set_cell(c, cert_data.get('Modelo y/o referencia tipo ', ''))
        print(f"  ✔ F.3.2.2: {cert_data.get('Modelo y/o referencia tipo ','')}")

    # ─────────────────────────────────────────────────────────────────────────
    # F.3.3.4  Tensión nominal (row 1) + Frecuencia nominal (row 2)
    # Both rows have clause "F.3.3.4"
    # ─────────────────────────────────────────────────────────────────────────
    f334_rows = [tr for tr in clause_trs
                 if (cells := tr.findall(f'{W}tc')) and
                 cell_text(cells[0]) == 'F.3.3.4']
    if len(f334_rows) >= 1:
        c = obs(f334_rows[0]); 
        if c is not None: set_cell(c, volt)
        print(f"  ✔ F.3.3.4 tension: {volt}")
    if len(f334_rows) >= 2:
        c = obs(f334_rows[1]); 
        if c is not None: set_cell(c, freq)
        print(f"  ✔ F.3.3.4 frecuencia: {freq}")

    # ─────────────────────────────────────────────────────────────────────────
    # F.3.3.6  Corriente / potencia
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('F.3.3.6')
    if tr is not None:
        c = obs(tr); 
        if c is not None: set_cell(c, amp)
        print(f"  ✔ F.3.3.6: {amp}")

    # ─────────────────────────────────────────────────────────────────────────
    # F.4  Instrucciones
    # Merged header row (3 cells) → veredicto in cell[2]
    # First row below → obs + veredicto
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('F.4')
    if tr is not None:
        # Header row veredicto
        c = vrd(tr); 
        if c is not None: set_cell(c, 'N' if is_psu else 'P')
        # First row below
        tr2 = _next(tr)
        if tr2 is not None:
            obs_txt = 'Instrucción en equipo a alimentar' if is_psu else 'posee manual'
            c = obs(tr2); 
            if c is not None: clear_cell(c); set_cell(c, obs_txt)
            c = vrd(tr2); 
            if c is not None: set_cell(c, 'N' if is_psu else 'P')
        print(f"  ✔ F.4: {'N' if is_psu else 'P'}")

    # ─────────────────────────────────────────────────────────────────────────
    # G.2  Relés  (merged header)
    # G.2.1 data row
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('G.2')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if ex(components,'rele') else 'N')
    tr = _find('G.2.1')
    if tr is not None:
        if ex(components, 'rele'):
            c = obs(tr); 
            if c is not None: set_cell(c, 'Rele certificado.(Ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
    print(f"  ✔ G.2/G.2.1: {'P' if ex(components,'rele') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # G.3.4  Fusible
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('G.3.4')
    if tr is not None:
        if ex(components, 'fusible'):
            c = obs(tr); 
            if c is not None: set_cell(c, 'Fusible certificado (ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = obs(tr); 
            if c is not None: clear_cell(c)
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ G.3.4: {'P' if ex(components,'fusible') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # G.6  Aislación de alambre  (merged header)
    # G.6.1  first data row
    # ─────────────────────────────────────────────────────────────────────────
    ais_ex   = ex(components,   'aislación del transformador')
    ais_cert = cert(components, 'aislación del transformador')

    tr = _find('G.6')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if ais_ex else 'N')
        # First row below G.6 header
        tr2 = _next(tr)
        if tr2 and cell_text(tr2.findall(f'{W}tc')[0]) == 'G.6.1':
            pass  # handled below
        elif tr2:
            # If first row below is not G.6.1 but a continuation row
            pass

    tr = _find('G.6.1')
    if tr is not None:
        if ais_ex and ais_cert:
            c = obs(tr); 
            if c is not None: clear_cell(c); set_cell(c, 'Aislacion del transformador certificado. (Ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        elif ais_ex and not ais_cert:
            c = obs(tr); 
            if c is not None: set_cell(c, '---')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
        else:
            # Transformer exists but no separate insulation → Anexo J text
            long = ('Transformador Cumple con Anexo J. Se realizó el ensayo J.2.2.1.2 a 6kVrms.  '
                    'El mandril usado fue de 4,0mm Los ensayos J.2.2.2, J.2.3 y J.2.4 se realizaron '
                    'a 3kVrms. La temperatura en el punto J.2.4 fue de 200ºC.')
            c = obs(tr); 
            if c is not None: clear_cell(c); set_cell(c, long)
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P' if ex(components,'transformador') else 'N')
    print(f"  ✔ G.6/G.6.1")

    # ─────────────────────────────────────────────────────────────────────────
    # G.8  Varistores  (merged header)
    # G.8.1 data row
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('G.8')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if ex(components,'varistor') else 'N')
    tr = _find('G.8.1')
    if tr is not None:
        if ex(components, 'varistor'):
            c = obs(tr); 
            if c is not None: clear_cell(c); set_cell(c, 'Varistor certificado.(Ver tabla 4.1.2)')
            c = vrd(tr); 
            if c is not None: set_cell(c, 'P')
        else:
            c = obs(tr); 
            if c is not None: clear_cell(c)
            c = vrd(tr); 
            if c is not None: set_cell(c, 'N')
    print(f"  ✔ G.8/G.8.1: {'P' if ex(components,'varistor') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # G.11  Capacitor y unidades RC  (merged header)
    # G.11.1 data row
    # ─────────────────────────────────────────────────────────────────────────
    has_cap = ex(components,'capacitor x') or ex(components,'capacitor y')
    tr = _find('G.11')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if has_cap else 'N')
    tr = _find('G.11.1')
    if tr is not None:
        combined = ''
        if ex(components, 'capacitor y'):
            combined = 'Capacitor Y certificado entre primario y secundario con aislación Y1. (ver tabla 4.1.2)'
        if ex(components, 'capacitor x'):
            add = 'Capacitor X certificado entre polos de alimentación con aislación X2. (ver tabla 4.1.2)'
            combined = (combined + '\n' + add).strip('\n') if combined else add
        c = obs(tr)
        if c and combined:
            set_cell(c, combined)
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if has_cap else 'N')
    print(f"  ✔ G.11/G.11.1: {'P' if has_cap else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # G.12  Optoacopladores
    # Row 0: merged header (3 cells) → veredicto in cell[2]
    # Row 1: data row (4 cells)
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('G.12')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if ex(components,'optoacoplador') else 'N')
        tr2 = _next(tr)
        if tr2 is not None:
            if ex(components, 'optoacoplador'):
                c = obs(tr2); 
                if c is not None: clear_cell(c); set_cell(c, 'Optoacoplador certificado. (Ver tabla 4.1.2)')
                c = vrd(tr2); 
                if c is not None: set_cell(c, 'P')
            else:
                c = obs(tr2); 
                if c is not None: clear_cell(c)
                c = vrd(tr2); 
                if c is not None: set_cell(c, 'N')
    print(f"  ✔ G.12: {'P' if ex(components,'optoacoplador') else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # P.1  Requisitos generales
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('P.1')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P' if is_sealed else 'N')
        print(f"  ✔ P.1: {'P' if is_sealed else 'N'}")

    # ─────────────────────────────────────────────────────────────────────────
    # P.2.2  Protección contra la entrada de objetos extraños
    # Row 0: veredicto P
    # Row 1: obs = "Equipo sellado" or alternative text
    # ─────────────────────────────────────────────────────────────────────────
    tr = _find('P.2.2')
    if tr is not None:
        c = vrd(tr); 
        if c is not None: set_cell(c, 'P')
        tr2 = _next(tr)
        if tr2 is not None:
            obs_txt = 'Equipo sellado' if is_sealed else \
                      'Partes con tensiones peligrosas fuera del área proyectada por las aberturas del equipo.'
            c = obs(tr2); 
            if c is not None: clear_cell(c); set_cell(c, obs_txt)
        print(f"  ✔ P.2.2: P, '{obs_txt}'")

    # ─────────────────────────────────────────────────────────────────────────
    # S.1  Multi-row block
    # Collect all S.1 rows within the main clause table
    # obs rows 1,2,3 → "(ver tabla S.1)" or "---"
    # vrd row 1 → P or N
    # vrd rows 4,5,6,7 → P or N
    # ─────────────────────────────────────────────────────────────────────────
    s1_block = []
    for i, tr_s in enumerate(clause_trs):
        cells = tr_s.findall(f'{W}tc')
        if not cells: continue
        ft = cell_text(cells[0])
        if ft == 'S.1':
            s1_block.append(tr_s)
            for j in range(i + 1, len(clause_trs)):
                nc = clause_trs[j].findall(f'{W}tc')
                if not nc: break
                nft = cell_text(nc[0])
                if nft and not nft.startswith('S.'):
                    break
                if nft != 'S':
                    s1_block.append(clause_trs[j])
            break

    def s1_obs(n):
        if n > len(s1_block): return None
        c = s1_block[n-1].findall(f'{W}tc')
        return c[2] if len(c) > 2 else None

    def s1_vrd(n):
        if n > len(s1_block): return None
        c = s1_block[n-1].findall(f'{W}tc')
        return c[3] if len(c) > 3 else None

    if clause_6482_P:
        for i in [1, 2, 3]:
            c = s1_obs(i)
            if c is not None: clear_cell(c); set_cell(c, ' (ver tabla S.1)')
        c = s1_vrd(1); 
        if c is not None: set_cell(c, 'P')
        for i in [4, 5, 6, 7]:
            c = s1_vrd(i); 
            if c is not None: set_cell(c, 'P')
        print(f"  ✔ S.1: P ({len(s1_block)} rows)")
    else:
        c = s1_obs(1); 
        if c is not None: clear_cell(c)
        for i in [2, 3]:
            c = s1_obs(i)
            if c is not None: clear_cell(c); set_cell(c, '---')
        c = s1_vrd(1); 
        if c is not None: set_cell(c, 'N')
        for i in [4, 5, 6, 7]:
            c = s1_vrd(i); 
            if c is not None: set_cell(c, 'N')
        print(f"  ✔ S.1: N ({len(s1_block)} rows)")

    # ── Save and pack ─────────────────────────────────────────────────────────
    doc_tree.write(str(doc_path),
                   xml_declaration=True, encoding='UTF-8', standalone=True)

    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(tmp.rglob('*')):
            if f.is_file():
                zf.write(f, f.relative_to(tmp))

    shutil.rmtree(tmp)
    print(f"\n  ✔ Output: {output_path}")
    return output_path


# ═══════════════════════════════════════════════════════════════════════════════
# CLI
# ═══════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder where the input .docx is located
    # docx_folder   = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where eut.json, certificadora.json and components.json are located
    # json_folder   = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where the output .docx is written before replacing the input
    # output_folder = docx_folder / 'output'
    # ── End of configuration ──────────────────────────────────────────────────

    output_folder.mkdir(parents=True, exist_ok=True)

    # Find input .docx (prefer "mod" in name, fall back to any .docx)
    mod_candidates = [f for f in docx_folder.iterdir()
                      if f.is_file() and f.suffix.lower() == '.docx'
                      and 'mod' in f.stem.lower()]
    any_candidates = [f for f in docx_folder.iterdir()
                      if f.is_file() and f.suffix.lower() == '.docx']

    if mod_candidates:
        docx_file = mod_candidates[0]
    elif any_candidates:
        docx_file = any_candidates[0]
    else:
        print(f"ERROR: No .docx found in: {docx_folder}"); sys.exit(1)

    # Output filename: ensure "mod" present, append "_cla"
    stem = docx_file.stem
    if 'mod' not in stem.lower():
        stem = stem + '_mod'
    output_file = output_folder / (stem + '_cla.docx')

    # Load JSON files
    missing = [f for f in ['eut.json', 'certificadora.json', 'components.json']
               if not (json_folder / f).exists()]
    if missing:
        print(f"ERROR: Missing files in '{json_folder}': {', '.join(missing)}")
        sys.exit(1)

    with open(json_folder / 'eut.json',            encoding='utf-8') as f:
        eut_data = json.load(f)
    with open(json_folder / 'certificadora.json',  encoding='utf-8') as f:
        cert_data = json.load(f)
    with open(json_folder / 'components.json',     encoding='utf-8') as f:
        components = json.load(f)

    print(f"\n  Input  : {docx_file.name}")
    print(f"  Output : {output_file.name}\n")

    apply_clauses(docx_file, eut_data, cert_data, components, output_file)

    # Move output to input_folder, delete original
    final_file = docx_folder / output_file.name
    docx_file.unlink()
    output_file.rename(final_file)
    print(f"  Replaced: {final_file.name}")
    print("\nDone.\n")
