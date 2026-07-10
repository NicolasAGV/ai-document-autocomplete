"""
docx_utils.py

Shared XML helpers for navigating OOXML table structure.
Imported by update_tabla_componentes.py and every column-filler module.
"""

import re


def find_row_boundaries(content: str) -> tuple[list[int], list[int]]:
    """Return parallel (starts, ends) character positions for every <w:tr>."""
    starts = [m.start() for m in re.finditer(r"<w:tr[ >]", content)]

    def closing(pos: int) -> int:
        depth, i = 0, pos
        while i < len(content):
            if content[i:i+5] == "<w:tr" and content[i+5:i+6] in (" ", ">"):
                depth += 1
                i += 5
            elif content[i:i+6] == "</w:tr" and content[i+6:i+7] in (" ", ">"):
                depth -= 1
                if depth == 0:
                    return content.index(">", i) + 1
                i += 6
            else:
                i += 1
        raise ValueError(f"No closing </w:tr> found starting at position {pos}.")

    return starts, [closing(s) for s in starts]


def find_top_level_row_boundaries(table_xml: str) -> tuple[list[int], list[int]]:
    """
    Return row positions for <w:tr> elements that are DIRECT children of the
    outermost <w:tbl>, skipping rows that belong to nested tables inside cells.

    Use this instead of find_row_boundaries when working on a single table's
    XML so that nested-table rows do not corrupt row indices.
    """
    tbl_depth = 0
    row_starts = []
    i = 0
    while i < len(table_xml):
        if table_xml[i:i+6] == "<w:tbl" and table_xml[i+6:i+7] in (" ", ">"):
            tbl_depth += 1
            i += 6
        elif table_xml[i:i+8] == "</w:tbl>":
            tbl_depth -= 1
            i += 8
        elif table_xml[i:i+5] == "<w:tr" and table_xml[i+5:i+6] in (" ", ">"):
            if tbl_depth == 1:  # direct child of the outermost table
                row_starts.append(i)
            i += 5
        else:
            i += 1

    def closing(pos: int) -> int:
        depth, i = 0, pos
        while i < len(table_xml):
            if table_xml[i:i+5] == "<w:tr" and table_xml[i+5:i+6] in (" ", ">"):
                depth += 1
                i += 5
            elif table_xml[i:i+6] == "</w:tr" and table_xml[i+6:i+7] in (" ", ">"):
                depth -= 1
                if depth == 0:
                    return table_xml.index(">", i) + 1
                i += 6
            else:
                i += 1
        raise ValueError(f"No closing </w:tr> found starting at position {pos}.")

    return row_starts, [closing(s) for s in row_starts]


def find_table_boundaries(content: str) -> tuple[list[int], list[int]]:
    """Return parallel (starts, ends) character positions for every <w:tbl>."""
    starts = [m.start() for m in re.finditer(r"<w:tbl[ >]", content)]

    def closing(pos: int) -> int:
        depth, i = 0, pos
        while i < len(content):
            if content[i:i+6] == "<w:tbl" and content[i+6:i+7] in (" ", ">"):
                depth += 1
                i += 6
            elif content[i:i+8] == "</w:tbl>":
                depth -= 1
                if depth == 0:
                    return i + 8
                i += 8
            else:
                i += 1
        raise ValueError(f"No closing </w:tbl> found starting at position {pos}.")

    return starts, [closing(s) for s in starts]


def find_table_by_content(content: str, text: str) -> tuple[int, int]:
    """
    Return (start, end) of the INNERMOST <w:tbl> whose visible text contains
    *text*.

    When a matching table is nested inside a wrapper/layout table, both tables
    contain the text; returning the smallest span gives us the most specific
    (innermost) table, which is the actual data table we want to edit.
    """
    starts, ends = find_table_boundaries(content)
    matches = []
    for s, e in zip(starts, ends):
        visible = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", content[s:e]))
        if text.lower() in visible.lower():
            matches.append((e - s, s, e))   # (size, start, end)
    if not matches:
        raise ValueError(f"No <w:tbl> containing the text {text!r} found in document.")
    matches.sort()          # smallest span first → innermost table
    _, s, e = matches[0]
    return s, e


def detect_data_rows(table_xml: str, n_cols: int = 6) -> tuple[int, list[int], list[int]]:
    """
    Return (first_data_idx, starts, ends) for a table whose header row has
    exactly *n_cols* cells.

    Uses find_top_level_row_boundaries so nested-table rows are ignored.
    first_data_idx is the index of the first data row (one past the header row).
    """
    starts, ends = find_top_level_row_boundaries(table_xml)
    for i, (s, e) in enumerate(zip(starts, ends)):
        col_starts, _ = find_cell_boundaries(table_xml[s:e])
        if len(col_starts) == n_cols:
            return i + 1, starts, ends
    raise RuntimeError(
        f"Could not find a row with {n_cols} cells in the table. "
        "Check that the correct table is being located."
    )


def find_cell_boundaries(row_xml: str) -> tuple[list[int], list[int]]:
    """Return (starts, ends) character positions for each <w:tc> in *row_xml*."""
    starts = [m.start() for m in re.finditer(r"<w:tc[ >]", row_xml)]

    def closing(pos: int) -> int:
        depth, i = 0, pos
        while i < len(row_xml):
            # <w:tc> or <w:tc ...> but NOT <w:tcPr>
            if row_xml[i:i+5] == "<w:tc" and row_xml[i+5:i+6] in (" ", ">"):
                depth += 1
                i += 5
            # </w:tc> is exactly 7 chars — also excludes </w:tcPr>
            elif row_xml[i:i+7] == "</w:tc>":
                depth -= 1
                if depth == 0:
                    return i + 7
                i += 7
            else:
                i += 1
        raise ValueError(f"No closing </w:tc> at position {pos}.")

    return starts, [closing(s) for s in starts]


def find_header_column(header_row_xml: str, text: str) -> int | None:
    """Return the 0-based column index of the header cell whose text contains *text*."""
    c_starts, c_ends = find_cell_boundaries(header_row_xml)
    for i, (s, e) in enumerate(zip(c_starts, c_ends)):
        cell_text = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", header_row_xml[s:e]))
        if text.lower() in cell_text.lower():
            return i
    return None
