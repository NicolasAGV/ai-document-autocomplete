"""
doc_mod_tabla_espigas.py
========================
Deletes BOTH "espigas de alimentación" tables (clause 4.7) from the report
docx when the EUT connection type is "Bornera".

The section is split into two consecutive tables:
  Table 1 — header:     clause "4.7" | "TABLA: Medición de espigas ..."
  Table 2 — data rows:  "Medición de:" | "Medida (mm)" | "Límite (mm)"

Logic:
    Read extracted_xlsx_eut_basic.json.
    If "Conexión de alimentación" == "Bornera"  →  remove both tables.
    Otherwise                                    →  leave the docx untouched.
"""

import json
import os
import re
import shutil
import sys
import tempfile
import zipfile
import importlib.util as _ilu
from pathlib import Path

# ── Paths ──────────────────────────────────────────────────────────────────────

sys.path.insert(0, str(Path(__file__).parent.parent))
import doc_mod_01_creation

docx_folder = doc_mod_01_creation.output_folder
json_folder = doc_mod_01_creation.json_folder

# ── Import tablas_finder (and its re-exported docx_utils helpers) ──────────────

_tf_path = Path(__file__).parent / "Varios general" / "tablas_finder.py"
_tf_spec = _ilu.spec_from_file_location("tablas_finder", _tf_path)
_tf_mod  = _ilu.module_from_spec(_tf_spec)   # type: ignore[arg-type]
_tf_spec.loader.exec_module(_tf_mod)          # type: ignore[union-attr]

find_table = _tf_mod.find_table

# ── Constants ──────────────────────────────────────────────────────────────────

TABLE_CLAUSE = "4.7"
TABLE_TITLE  = "TABLA: Medición de espigas de alimentación"

CONEXION_KEY   = "Conexión de alimentación"
CONEXION_VALUE = "Bornera"

# ── Helpers ────────────────────────────────────────────────────────────────────

def read_docx_xml(docx_path: Path) -> str:
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open("word/document.xml") as f:
            return f.read().decode("utf-8")


def write_docx_xml(docx_path: Path, new_content: str,
                   output_path: Path | None = None) -> Path:
    if output_path is None:
        output_path = docx_path
    fd, tmp_str = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    tmp = Path(tmp_str)
    with zipfile.ZipFile(docx_path) as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, new_content.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
    shutil.move(str(tmp), str(output_path))
    return output_path


def delete_two_tables(content: str, s1: int, e1: int, s2: int, e2: int) -> str:
    """Remove both table XML blocks in a single splice (requires s2 >= e1)."""
    return content[:s1] + content[e1:s2] + content[e2:]


def _build_two_tables_xml(content: str, t1_xml: str, t2_xml: str) -> str:
    """Return *content* with the body stripped to contain only *t1_xml* + *t2_xml*."""
    body_start   = content.find("<w:body")
    body_tag_end = content.index(">", body_start) + 1
    body_close   = content.rfind("</w:body>")
    body_inner   = content[body_tag_end:body_close]
    sect_match   = re.search(r"<w:sectPr\b[\s\S]*?</w:sectPr>", body_inner)
    sect_pr      = sect_match.group(0) if sect_match else ""
    return content[:body_tag_end] + t1_xml + t2_xml + sect_pr + content[body_close:]


def _apply_strikethrough(xml: str) -> str:
    """Add <w:strike/> to every run inside *xml*."""
    def patch_rpr(m: re.Match) -> str:
        rpr = m.group(0)
        if "<w:strike" not in rpr:
            close = rpr.index(">") + 1
            return rpr[:close] + "<w:strike/>" + rpr[close:]
        return rpr

    def patch_run(m: re.Match) -> str:
        run = m.group(0)
        if "<w:rPr" not in run:
            close = run.index(">") + 1
            return run[:close] + "<w:rPr><w:strike/></w:rPr>" + run[close:]
        return run

    xml = re.sub(r"<w:rPr\b[^>]*>[\s\S]*?</w:rPr>", patch_rpr, xml)
    xml = re.sub(r"<w:r\b[^>]*>[\s\S]*?</w:r>", patch_run, xml)
    return xml


# ── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    # Load EUT basic info
    eut_basic_path = json_folder / doc_mod_01_creation.EUT_basic_json
    if not eut_basic_path.exists():
        print(f"ERROR: '{eut_basic_path}' not found.")
        raise SystemExit(1)

    with open(eut_basic_path, encoding="utf-8") as f:
        eut_basic = json.load(f)

    conexion = eut_basic.get(CONEXION_KEY, "").strip()
    print(f"  {CONEXION_KEY}: '{conexion}'")

    if conexion.lower() != CONEXION_VALUE.lower():
        print(f"  → Not '{CONEXION_VALUE}'. Tables kept as-is.")
        print("\nDone.\n")
        return

    # Resolve the docx file
    informe_json = json_folder / "doc_mod_numero_informe.json"
    if not informe_json.exists():
        print(f"ERROR: '{informe_json}' not found. Run doc_mod_01_creation.py first.")
        raise SystemExit(1)

    with open(informe_json, encoding="utf-8") as f:
        info = json.load(f)

    docx_file = docx_folder / info.get("filename", "")
    if not docx_file.exists():
        print(f"ERROR: docx not found: '{docx_file}'")
        raise SystemExit(1)

    print(f"  Docx : {docx_file.name}")
    content = read_docx_xml(docx_file)

    # Locate Table 1 (clause/title header)
    try:
        s1, e1 = find_table(content, clause=TABLE_CLAUSE, title=TABLE_TITLE)
    except ValueError as exc:
        print(f"  WARNING: header table not found — {exc}")
        print("\nDone.\n")
        return
    print(f"  ✔ Header table found  (XML pos {s1}–{e1})")

    # Locate Table 2 (data rows: Medición de / Medida / Límite)
    try:
        s2, e2 = find_table(content, title="Medida (mm)")
    except ValueError as exc:
        print(f"  WARNING: data table not found — {exc}")
        print("\nDone.\n")
        return
    print(f"  ✔ Data table found    (XML pos {s2}–{e2})")

    # Save secondary docx with only the two tables, text struck through
    struck_t1 = _apply_strikethrough(content[s1:e1])
    struck_t2 = _apply_strikethrough(content[s2:e2])
    struck_content = _build_two_tables_xml(content, struck_t1, struck_t2)
    struck_folder = docx_folder / ".docx"
    struck_folder.mkdir(parents=True, exist_ok=True)
    struck_path = struck_folder / "doc_mod_espigas.docx"
    write_docx_xml(docx_file, struck_content, output_path=struck_path)
    print(f"  ✔ Strikethrough docx saved: {struck_path}")

    # Delete both in a single splice and save
    new_content = delete_two_tables(content, s1, e1, s2, e2)
    write_docx_xml(docx_file, new_content)
    print(f"  ✔ Both tables deleted and docx saved: {docx_file.name}")
    print("\nDone.\n")


if __name__ == "__main__":
    main()
