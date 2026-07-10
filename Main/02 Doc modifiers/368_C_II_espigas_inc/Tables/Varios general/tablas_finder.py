"""
tablas_finder.py — Locate tables in a .docx by clause number and/or title.

Tables are identified by the content of their FIRST ROW:
  - clause : text expected in the FIRST CELL  (e.g. "4.7" or "5.4.1.4,")
  - title  : text expected in ANY CELL        (e.g. "Medición de espigas")

Both parameters are optional but at least one must be supplied.
Matching is case-insensitive and substring-based, so partial strings work.

Exposes:
    find_table(content, clause, title)   → (start, end)  XML char positions
    list_table_headers(content)          → [(clause_cell, rest_cells), ...]
    read_docx_xml(path)                  → document.xml content as str
"""

import re
import sys
import zipfile
import importlib.util as _ilu
from pathlib import Path

# ── Import shared XML helpers from the sibling utility module ─────────────────

_du_path = Path(__file__).parent.parent / "Varios tabla cond ensayo" / "docx_utils.py"
_du_spec = _ilu.spec_from_file_location("docx_utils", _du_path)
_du_mod  = _ilu.module_from_spec(_du_spec)   # type: ignore[arg-type]
_du_spec.loader.exec_module(_du_mod)          # type: ignore[union-attr]

find_table_boundaries         = _du_mod.find_table_boundaries
find_top_level_row_boundaries = _du_mod.find_top_level_row_boundaries
find_cell_boundaries          = _du_mod.find_cell_boundaries


# ── Internal helpers ──────────────────────────────────────────────────────────

def _visible_text(xml: str) -> str:
    """Concatenate all <w:t> text runs in *xml* into a single string."""
    return "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", xml)).strip()


def _first_row_cell_texts(table_xml: str) -> list[str]:
    """Return visible text for each cell in the first (top-level) row."""
    row_starts, row_ends = find_top_level_row_boundaries(table_xml)
    if not row_starts:
        return []
    first_row = table_xml[row_starts[0]:row_ends[0]]
    c_starts, c_ends = find_cell_boundaries(first_row)
    return [_visible_text(first_row[s:e]) for s, e in zip(c_starts, c_ends)]


def _table_matches(table_xml: str, clause: str | None, title: str | None) -> bool:
    """Return True when the table's first row satisfies both filter criteria."""
    cells = _first_row_cell_texts(table_xml)
    if not cells:
        return False

    if clause is not None:
        # Clause is expected in the first cell (normalised, case-insensitive)
        if clause.strip().lower() not in cells[0].lower():
            return False

    if title is not None:
        # Title may appear in any cell of the first row
        title_norm = title.strip().lower()
        if not any(title_norm in c.lower() for c in cells):
            return False

    return True


# ── Public API ────────────────────────────────────────────────────────────────

def find_table(
    content: str,
    clause: str | None = None,
    title:  str | None = None,
) -> tuple[int, int]:
    """
    Return (start, end) XML character positions of the table whose first row
    contains *clause* in its first cell and/or *title* in any cell.

    When a matching data table is nested inside a layout/wrapper table, both
    share the same text; the innermost (smallest-span) match is returned.

    Raises ValueError  — no match found, or no search criteria given.
    Raises RuntimeError — ambiguous (multiple equally-sized matches).
    """
    if clause is None and title is None:
        raise ValueError("Supply at least one of 'clause' or 'title'.")

    t_starts, t_ends = find_table_boundaries(content)

    candidates: list[tuple[int, int, int]] = []   # (span_size, start, end)
    for s, e in zip(t_starts, t_ends):
        if _table_matches(content[s:e], clause, title):
            candidates.append((e - s, s, e))

    if not candidates:
        parts: list[str] = []
        if clause:
            parts.append(f"clause={clause!r}")
        if title:
            parts.append(f"title={title!r}")
        raise ValueError(f"No table found matching {', '.join(parts)}.")

    # Sort by span so the innermost (smallest) candidate comes first
    candidates.sort()
    _, s, e = candidates[0]
    return s, e


def list_table_headers(content: str) -> list[tuple[str, str]]:
    """
    Return (first_cell, remaining_cells_joined) for every table's first row.

    Useful for discovering the exact clause and title strings present in a
    document before calling find_table().
    """
    t_starts, t_ends = find_table_boundaries(content)
    result: list[tuple[str, str]] = []
    for s, e in zip(t_starts, t_ends):
        cells = _first_row_cell_texts(content[s:e])
        if cells:
            result.append((cells[0], " | ".join(cells[1:])))
    return result


def read_docx_xml(docx_path: Path) -> str:
    """Return the word/document.xml content of *docx_path* as a UTF-8 string."""
    with zipfile.ZipFile(docx_path) as zf:
        with zf.open("word/document.xml") as f:
            return f.read().decode("utf-8")


# ── Standalone demo ───────────────────────────────────────────────────────────

def main() -> None:
    import json
    sys.path.insert(0, str(Path(__file__).parent.parent))
    import doc_mod_01_creation

    informe_json = doc_mod_01_creation.json_folder / "doc_mod_numero_informe.json"
    if not informe_json.exists():
        print(f"ERROR: '{informe_json}' not found. Run doc_mod_01_creation.py first.")
        raise SystemExit(1)

    with open(informe_json, encoding="utf-8") as f:
        info = json.load(f)

    docx_file = doc_mod_01_creation.output_folder / info.get("filename", "")
    if not docx_file.exists():
        print(f"ERROR: docx not found: '{docx_file}'")
        raise SystemExit(1)

    print(f"  Docx : {docx_file.name}\n")
    content = read_docx_xml(docx_file)

    # ── List all table headers found in the document ──────────────────────────
    headers = list_table_headers(content)
    print(f"  Tables found: {len(headers)}")
    for i, (clause_cell, rest) in enumerate(headers):
        print(f"    [{i:02d}]  clause='{clause_cell}'  |  {rest}")

    # ── Example searches ──────────────────────────────────────────────────────
    examples = [
        dict(clause="4.7",     title="Medición de espigas de alimentación"),
        dict(clause="5.4.1.4", title="Mediciones de temperatura"),
        dict(clause=None,      title="Mediciones de temperatura"),
        dict(clause="4.7",     title=None),
    ]

    print()
    for kwargs in examples:
        label_parts = [f"{k}={v!r}" for k, v in kwargs.items() if v is not None]
        label = ", ".join(label_parts)
        try:
            start, end = find_table(content, **kwargs)
            print(f"  ✔  find_table({label})")
            print(f"       XML pos {start}–{end}  ({end - start} chars)")
        except ValueError as exc:
            print(f"  ✘  find_table({label})")
            print(f"       {exc}")

    print("\nDone.\n")


if __name__ == "__main__":
    main()
