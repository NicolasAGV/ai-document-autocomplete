"""
tabla_condiciones_ensayo.py  —  Locates the "Condiciones de Ensayo" table and
fills empty TEMP (°C) and HUME-DAD (%) cells with random values within
configurable ranges.

Header columns (left to right):
    CLAUSULA | TABLA | INL. | INCERTI-DUMBRE (%) | TEMP (°C) | HUME-DAD (%) |
    INSTRUMENTOS UTILIZADOS | COMENTARIOS | FECHA

Exposes:
    find_table(content)                → (start, end) raw XML positions of the table
    fill_temp_column(content, ...)     → (modified_content, cells_filled)
    fill_hume_column(content, ...)     → (modified_content, cells_filled)
    read_docx_xml(path)                → document.xml content as str
    write_docx_xml(path, content, ...) → saves modified XML back into the docx
    N_COLS                             → expected column count (9)
    HEADER_COLUMNS                     → ordered list of column-name substrings
"""

import os
import re
import sys
import json
import random
import shutil
import zipfile
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import doc_mod_01_creation

import importlib.util as _ilu

_du_path = Path(__file__).parent / "Varios tabla cond ensayo" / "docx_utils.py"
_du_spec = _ilu.spec_from_file_location("docx_utils", _du_path)
_du_mod  = _ilu.module_from_spec(_du_spec)   # type: ignore[arg-type]
_du_spec.loader.exec_module(_du_mod)          # type: ignore[union-attr]

find_table_by_content         = _du_mod.find_table_by_content
find_top_level_row_boundaries = _du_mod.find_top_level_row_boundaries
find_cell_boundaries          = _du_mod.find_cell_boundaries

# ── Fill parameters — edit these values to change the random ranges ────────────

TEMP_MIN: float = 22.0   # °C  lower bound for empty TEMP cells
TEMP_MAX: float = 25.9   # °C  upper bound for empty TEMP cells
TEMP_DECIMAL_SEP: str = ","  # decimal separator for TEMP values

HUME_MIN: float = 36.0   # %RH  lower bound for empty HUME-DAD cells
HUME_MAX: float = 44.9   # %RH  upper bound for empty HUME-DAD cells
HUME_DECIMAL_SEP: str = ","  # decimal separator for HUME-DAD values

# ── Table identity ─────────────────────────────────────────────────────────────

TABLE_ANCHOR = "INSTRUMENTOS UTILIZADOS"

HEADER_COLUMNS = [
    "CLAUSULA",
    "TABLA",
    "INL",
    "INCERTI",      # INCERTI-DUMBRE (%)
    "TEMP",         # TEMP (°C)
    "HUME",         # HUME-DAD (%)
    "INSTRUMENTOS", # INSTRUMENTOS UTILIZADOS
    "COMENTARIOS",
    "FECHA",
]

N_COLS = len(HEADER_COLUMNS)

# ── Path configuration ─────────────────────────────────────────────────────────

docx_folder = doc_mod_01_creation.output_folder
json_folder = doc_mod_01_creation.json_folder

# ── Internal helpers ───────────────────────────────────────────────────────────

def _cell_visible_text(cell_xml: str) -> str:
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", cell_xml)).strip()


def _inject_text_into_cell(cell_xml: str, value: str) -> str:
    """Insert a plain run into the first <w:p> of *cell_xml*."""
    run = f'<w:r><w:t>{value}</w:t></w:r>'
    return cell_xml.replace('</w:p>', run + '</w:p>', 1)


# ── Public API ─────────────────────────────────────────────────────────────────

def find_table(content: str) -> tuple[int, int]:
    """Return (start, end) XML character positions of the Condiciones de Ensayo table."""
    return find_table_by_content(content, TABLE_ANCHOR)


def read_docx_xml(docx_path: Path) -> str:
    """Return the document.xml content of *docx_path* as a UTF-8 string."""
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open('word/document.xml') as f:
            return f.read().decode('utf-8')


def write_docx_xml(docx_path: Path, new_content: str,
                   output_path: Path | None = None) -> Path:
    """
    Write *new_content* as word/document.xml into the docx.
    Copies all other zip entries unchanged.
    If *output_path* is None the source file is overwritten in place.
    Returns the output path.
    """
    if output_path is None:
        output_path = docx_path

    fd, tmp_str = tempfile.mkstemp(suffix='.docx')
    os.close(fd)
    tmp = Path(tmp_str)
    with zipfile.ZipFile(docx_path) as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, new_content.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))

    shutil.move(str(tmp), str(output_path))
    return output_path


def _fill_column(
    content: str,
    col_header: str,
    val_min: float,
    val_max: float,
    decimal_sep: str,
) -> tuple[str, int]:
    """
    Fill every empty data cell in *col_header* column with a random value in
    [val_min, val_max], formatted to one decimal place using *decimal_sep*.

    Cells that already contain a value are left untouched.
    Rows are processed in reverse so earlier XML positions remain valid after
    each replacement.

    Returns (modified_content, number_of_cells_filled).
    """
    t_start, t_end = find_table(content)
    table_xml = content[t_start:t_end]

    row_starts, row_ends = find_top_level_row_boundaries(table_xml)
    if not row_starts:
        raise RuntimeError("Condiciones de Ensayo table has no rows.")

    # Identify the target column index from the header row
    header_xml = table_xml[row_starts[0]:row_ends[0]]
    c_starts_h, c_ends_h = find_cell_boundaries(header_xml)
    col_idx: int | None = None
    for i, (cs, ce) in enumerate(zip(c_starts_h, c_ends_h)):
        txt = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", header_xml[cs:ce]))
        if col_header.upper() in txt.upper():
            col_idx = i
            break

    if col_idx is None:
        raise RuntimeError(f"Column '{col_header}' not found in Condiciones de Ensayo table header.")

    modified_table = table_xml
    filled = 0

    for row_idx in range(len(row_starts) - 1, 0, -1):
        rs, re_ = row_starts[row_idx], row_ends[row_idx]
        row_xml = modified_table[rs:re_]

        c_starts, c_ends = find_cell_boundaries(row_xml)
        if col_idx >= len(c_starts):
            continue

        cs, ce = c_starts[col_idx], c_ends[col_idx]
        cell_xml = row_xml[cs:ce]

        if _cell_visible_text(cell_xml):
            continue   # already has a value — leave it untouched

        value    = f"{random.uniform(val_min, val_max):.1f}".replace(".", decimal_sep)
        new_cell = _inject_text_into_cell(cell_xml, value)
        new_row  = row_xml[:cs] + new_cell + row_xml[ce:]
        modified_table = modified_table[:rs] + new_row + modified_table[re_:]
        filled += 1

    return content[:t_start] + modified_table + content[t_end:], filled


def fill_temp_column(
    content: str,
    temp_min: float = TEMP_MIN,
    temp_max: float = TEMP_MAX,
    decimal_sep: str = TEMP_DECIMAL_SEP,
) -> tuple[str, int]:
    """Fill empty TEMP (°C) cells. Returns (modified_content, cells_filled)."""
    return _fill_column(content, "TEMP", temp_min, temp_max, decimal_sep)


def fill_hume_column(
    content: str,
    hume_min: float = HUME_MIN,
    hume_max: float = HUME_MAX,
    decimal_sep: str = HUME_DECIMAL_SEP,
) -> tuple[str, int]:
    """Fill empty HUME-DAD (%) cells. Returns (modified_content, cells_filled)."""
    return _fill_column(content, "HUME", hume_min, hume_max, decimal_sep)


# ── Standalone: locate table + fill TEMP column ────────────────────────────────

def main() -> None:
    informe_json = json_folder / "doc_mod_numero_informe.json"
    if not informe_json.exists():
        print(f"ERROR: '{informe_json}' not found. Run doc_mod_01_creation.py first.")
        raise SystemExit(1)

    with open(informe_json, encoding='utf-8') as f:
        info = json.load(f)

    docx_file = docx_folder / info.get("filename", "")
    if not docx_file.exists():
        print(f"ERROR: docx not found: '{docx_file}'")
        raise SystemExit(1)

    print(f"  Docx : {docx_file.name}")

    content = read_docx_xml(docx_file)

    # Locate and report table structure
    try:
        t_start, t_end = find_table(content)
    except ValueError as exc:
        print(f"  ⚠ Table not found: {exc}")
        raise SystemExit(1)

    table_xml = content[t_start:t_end]
    row_starts, row_ends = find_top_level_row_boundaries(table_xml)
    print(f"  ✔ Table found   (XML pos {t_start}–{t_end}, {len(table_xml)} chars)")
    print(f"  ✔ Rows          : {len(row_starts)}")

    if row_starts:
        header_xml = table_xml[row_starts[0]:row_ends[0]]
        c_starts, c_ends = find_cell_boundaries(header_xml)
        print(f"  ✔ Header cells  : {len(c_starts)}  (expected {N_COLS})")
        for i, (cs, ce) in enumerate(zip(c_starts, c_ends)):
            cell_text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", header_xml[cs:ce]))
            expected  = HEADER_COLUMNS[i] if i < N_COLS else "?"
            mark      = "✔" if expected.lower() in cell_text.lower() else "⚠"
            print(f"    {mark} col {i}: '{cell_text.strip()}'")

    # Fill empty TEMP and HUME-DAD cells
    print(f"\n  — Filling empty TEMP cells  [{TEMP_MIN}, {TEMP_MAX}] —")
    new_content, filled_temp = fill_temp_column(content)
    print(f"  ✔ Filled {filled_temp} empty cell(s)")

    print(f"\n  — Filling empty HUME-DAD cells  [{HUME_MIN}, {HUME_MAX}] —")
    new_content, filled_hume = fill_hume_column(new_content)
    print(f"  ✔ Filled {filled_hume} empty cell(s)")

    if filled_temp or filled_hume:
        write_docx_xml(docx_file, new_content)
        print(f"\n  ✔ Saved: {docx_file.name}")

    print("\nDone.\n")


if __name__ == "__main__":
    main()
