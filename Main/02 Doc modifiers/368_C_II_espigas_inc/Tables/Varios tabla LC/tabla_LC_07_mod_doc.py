"""
tabla_LC_mod_doc.py  —  step 2 of the pipeline.

Run after tabla_LC_nombre.py:
    python tabla_LC_mod_doc.py

What it does
------------
  1. Read the populated doc_mod_listado_componentes.json (filled by tabla_LC_nombre.py
     and the individual filler scripts).
  2. Locate table 4.1.2 in the .docx found in DOCX_FOLDER.
  3. Write component names into column 0 ("objeto/parte No.").
  4. Write or embed each non-empty slot value into its matching column.
     String values → text cell.
     File-path values → inline image (centred, aspect-ratio preserved).
  5. Save the result to OUTPUT_FOLDER.
"""

import re
import sys
import json
import shutil
import zipfile
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import doc_mod_01_creation

# =========================================================================== #
# CONFIGURATION                                                                #
# =========================================================================== #

MAIN_FOLDER   = doc_mod_01_creation.source_folder
DOCX_FOLDER   = MAIN_FOLDER / "output"
JSON_FOLDER   = doc_mod_01_creation.json_folder
OUTPUT_FOLDER = MAIN_FOLDER / "output"

# Maximum size for logo/certification images (longest side, centimetres).
LOGO_SIZE_CM: float = 1.0

# Width for slot [1] marcado images (tipo/modelo column), centimetres.
MODELO_SIZE_CM: float = 2.6

# Number of data columns after "objeto/parte No." (must match tabla_LC_nombre.py).
N_DATA_COLS: int = 5


# =========================================================================== #
# XML HELPERS                                                                  #
# =========================================================================== #

def _find_top_level_row_boundaries(table_xml: str) -> tuple[list[int], list[int]]:
    tbl_depth, row_starts, i = 0, [], 0
    while i < len(table_xml):
        if table_xml[i:i+6] == "<w:tbl" and table_xml[i+6:i+7] in (" ", ">"):
            tbl_depth += 1; i += 6
        elif table_xml[i:i+8] == "</w:tbl>":
            tbl_depth -= 1; i += 8
        elif table_xml[i:i+5] == "<w:tr" and table_xml[i+5:i+6] in (" ", ">"):
            if tbl_depth == 1:
                row_starts.append(i)
            i += 5
        else:
            i += 1

    def row_end(pos: int) -> int:
        depth, j = 0, pos
        while j < len(table_xml):
            if table_xml[j:j+5] == "<w:tr" and table_xml[j+5:j+6] in (" ", ">"):
                depth += 1; j += 5
            elif table_xml[j:j+6] == "</w:tr" and table_xml[j+6:j+7] in (" ", ">"):
                depth -= 1
                if depth == 0:
                    return table_xml.index(">", j) + 1
                j += 6
            else:
                j += 1
        raise ValueError(f"No closing </w:tr> from position {pos}.")

    return row_starts, [row_end(s) for s in row_starts]


def _find_cell_boundaries(row_xml: str) -> tuple[list[int], list[int]]:
    starts = [m.start() for m in re.finditer(r"<w:tc[ >]", row_xml)]

    def cell_end(pos: int) -> int:
        depth, j = 0, pos
        while j < len(row_xml):
            if row_xml[j:j+5] == "<w:tc" and row_xml[j+5:j+6] in (" ", ">"):
                depth += 1; j += 5
            elif row_xml[j:j+7] == "</w:tc>":
                depth -= 1
                if depth == 0:
                    return j + 7
                j += 7
            else:
                j += 1
        raise ValueError(f"No closing </w:tc> from position {pos}.")

    return starts, [cell_end(s) for s in starts]


def _find_table_boundaries(content: str) -> tuple[list[int], list[int]]:
    starts = [m.start() for m in re.finditer(r"<w:tbl[ >]", content)]

    def tbl_end(pos: int) -> int:
        depth, j = 0, pos
        while j < len(content):
            if content[j:j+6] == "<w:tbl" and content[j+6:j+7] in (" ", ">"):
                depth += 1; j += 6
            elif content[j:j+8] == "</w:tbl>":
                depth -= 1
                if depth == 0:
                    return j + 8
                j += 8
            else:
                j += 1
        raise ValueError(f"No closing </w:tbl> from position {pos}.")

    return starts, [tbl_end(s) for s in starts]


def _build_table_only_xml(xml_content: str, table_xml: str) -> str:
    """Return xml_content with the body stripped to contain only *table_xml*."""
    body_start   = xml_content.find("<w:body")
    body_tag_end = xml_content.index(">", body_start) + 1
    body_close   = xml_content.rfind("</w:body>")
    body_inner   = xml_content[body_tag_end:body_close]

    sect_match = re.search(r"<w:sectPr\b[\s\S]*?</w:sectPr>", body_inner)
    sect_pr    = sect_match.group(0) if sect_match else ""

    return xml_content[:body_tag_end] + table_xml + sect_pr + xml_content[body_close:]


def _find_table_by_content(content: str, text: str) -> tuple[int, int]:
    """Return the innermost <w:tbl> whose visible text contains *text*."""
    starts, ends = _find_table_boundaries(content)
    matches = []
    for s, e in zip(starts, ends):
        visible = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", content[s:e]))
        if text.lower() in visible.lower():
            matches.append((e - s, s, e))
    if not matches:
        raise ValueError(f"No table containing {text!r} found.")
    matches.sort()
    _, s, e = matches[0]
    return s, e


def _detect_data_rows(table_xml: str, n_cols: int = 6) -> tuple[int, list[int], list[int]]:
    """Return (first_data_idx, starts, ends); header row = first row with n_cols cells."""
    starts, ends = _find_top_level_row_boundaries(table_xml)
    for i, (s, e) in enumerate(zip(starts, ends)):
        cs, _ = _find_cell_boundaries(table_xml[s:e])
        if len(cs) == n_cols:
            return i + 1, starts, ends
    raise RuntimeError(f"No row with {n_cols} cells found — wrong table?")


# =========================================================================== #
# TEXT CELL HELPERS                                                             #
# =========================================================================== #

_HIGHLIGHT_RUN = (
    '<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr>'
    '<w:t xml:space="preserve">{}</w:t></w:r>'
)


def _set_cell_text(cell_xml: str, text: str) -> str:
    safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    run  = _HIGHLIGHT_RUN.format(safe)

    # Replace the whole <w:r>…</w:r> that owns the first <w:t>
    t_match = re.search(r"<w:t(?:\s[^>]*)?>([^<]*)</w:t>", cell_xml)
    if t_match:
        r_start = cell_xml.rfind("<w:r", 0, t_match.start())
        r_end   = cell_xml.find("</w:r>", t_match.end())
        if r_start != -1 and r_end != -1:
            return cell_xml[:r_start] + run + cell_xml[r_end + len("</w:r>"):]
        # <w:t> not wrapped in <w:r> — just replace the <w:t> itself
        return cell_xml[:t_match.start()] + run + cell_xml[t_match.end():]

    # No <w:t> at all — insert before </w:p>
    p = cell_xml.find("</w:p>")
    if p != -1:
        return cell_xml[:p] + run + cell_xml[p:]
    return cell_xml


def _set_text_in_row(row_xml: str, col_idx: int, text: str) -> str:
    cs, ce = _find_cell_boundaries(row_xml)
    if col_idx >= len(cs):
        return row_xml
    cell = row_xml[cs[col_idx]:ce[col_idx]]
    return row_xml[:cs[col_idx]] + _set_cell_text(cell, text) + row_xml[ce[col_idx]:]


# =========================================================================== #
# IMAGE CELL HELPERS                                                            #
# =========================================================================== #

_CM_TO_EMU = 360_000

_IMAGE_MIME = {
    "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "png": "image/png",  "bmp":  "image/bmp",
    "tiff": "image/tiff","tif":  "image/tiff",
}

_DRAWING_NS = {
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _next_r_id(rels_xml: str) -> str:
    existing = [int(m.group(1)) for m in re.finditer(r'Id="rId(\d+)"', rels_xml)]
    return f"rId{max(existing, default=0) + 1}"


def _add_rel(rels_xml: str, r_id: str, media_filename: str) -> str:
    entry = (
        f'<Relationship Id="{r_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" '
        f'Target="media/{media_filename}"/>'
    )
    idx = rels_xml.rfind("</Relationships>")
    return (rels_xml[:idx] + entry + rels_xml[idx:]) if idx != -1 else rels_xml + entry


def _ensure_ct(ct_xml: str, ext: str) -> str:
    ext_lower = ext.lstrip(".").lower()
    mime = _IMAGE_MIME.get(ext_lower, "application/octet-stream")
    if f'Extension="{ext_lower}"' not in ct_xml:
        entry = f'<Default Extension="{ext_lower}" ContentType="{mime}"/>'
        idx = ct_xml.rfind("</Types>")
        ct_xml = (ct_xml[:idx] + entry + ct_xml[idx:]) if idx != -1 else ct_xml + entry
    return ct_xml


def _ensure_drawing_namespaces(xml_content: str) -> str:
    for prefix, uri in _DRAWING_NS.items():
        if f"xmlns:{prefix}=" not in xml_content:
            xml_content = re.sub(
                r"(<w:document\b)", rf'\1 xmlns:{prefix}="{uri}"', xml_content, count=1
            )
    return xml_content


def _build_drawing(r_id: str, cx: int, cy: int, pic_id: int, label: str) -> str:
    safe = label.replace("&","&amp;").replace('"',"&quot;").replace("<","&lt;").replace(">","&gt;")
    return (
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{pic_id}" name="{safe}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:nvPicPr><pic:cNvPr id="{pic_id}" name="{safe}"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{r_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        f'</pic:pic></a:graphicData></a:graphic></wp:inline>'
    )





def _register_image(img_path: Path, tmp_dir: Path, tag: str) -> str:
    media_dir = tmp_dir / "word" / "media"
    media_dir.mkdir(exist_ok=True)
    media_name = f"{tag}{img_path.suffix.lower()}"
    shutil.copy2(img_path, media_dir / media_name)

    rels_path = tmp_dir / "word" / "_rels" / "document.xml.rels"
    rels_xml  = rels_path.read_text(encoding="utf-8")
    r_id      = _next_r_id(rels_xml)
    rels_path.write_text(_add_rel(rels_xml, r_id, media_name), encoding="utf-8")

    ct_path = tmp_dir / "[Content_Types].xml"
    ct_path.write_text(_ensure_ct(ct_path.read_text(encoding="utf-8"), img_path.suffix), encoding="utf-8")
    return r_id


def _image_size_emu(img_path: Path, max_cm: float) -> tuple[int, int]:
    """
    Return (cx, cy) in EMU scaled to fit within max_cm × max_cm while
    preserving the image's aspect ratio. Falls back to a square on failure.
    Supports JPEG and PNG without external dependencies.
    """
    import struct
    w_px = h_px = 0
    try:
        data = img_path.read_bytes()
        if data[:8] == b'\x89PNG\r\n\x1a\n':          # PNG
            w_px, h_px = struct.unpack('>II', data[16:24])
        elif data[:2] == b'\xff\xd8':                  # JPEG
            pos = 2
            while pos < len(data) - 4:
                if data[pos] != 0xFF:
                    break
                marker = data[pos + 1]
                if marker in (0xC0, 0xC1, 0xC2):
                    h_px, w_px = struct.unpack('>HH', data[pos + 5: pos + 9])
                    break
                seg_len = struct.unpack('>H', data[pos + 2: pos + 4])[0]
                pos += 2 + seg_len
    except Exception:
        pass

    max_emu = int(max_cm * _CM_TO_EMU)
    if w_px <= 0 or h_px <= 0:
        return max_emu, max_emu
    if w_px >= h_px:
        return max_emu, int(max_emu * h_px / w_px)
    return int(max_emu * w_px / h_px), max_emu


# =========================================================================== #
# FIRST-COLUMN FILL  ("objeto/parte No.")                                      #
# =========================================================================== #

def _fill_first_column(
    table_xml: str,
    component_names: list[str],
    first_data: int,
    starts: list[int],
    ends: list[int],
) -> str:
    patches: list[tuple[int, int, str]] = []
    for i, name in enumerate(component_names):
        row_idx = first_data + i
        if row_idx >= len(starts):
            print(f"  [col 0] WARNING: ran out of rows at component '{name}' — stopping.")
            break
        row_xml = table_xml[starts[row_idx]:ends[row_idx]]
        patches.append((starts[row_idx], ends[row_idx], _set_text_in_row(row_xml, 0, name)))

    for s, e, new in reversed(patches):
        table_xml = table_xml[:s] + new + table_xml[e:]
    print(f"  [col 0] Filled {len(patches)} row(s) with component names.")
    return table_xml


# =========================================================================== #
# CORE: apply one column's data                                                #
# =========================================================================== #

def apply_fills(
    table_xml: str,
    component_names: list[str],
    first_data: int,
    starts: list[int],
    ends: list[int],
    col_idx: int,
    data: dict[str, "str | Path"],
    tmp_dir: Path,
    img_max_cm: float = LOGO_SIZE_CM,
) -> str:
    img_cache: dict[Path, str] = {}
    patches: list[tuple[int, int, str]] = []

    for i, name in enumerate(component_names):
        row_idx = first_data + i
        if row_idx >= len(starts) or name not in data:
            continue
        value   = data[name]
        row_xml = table_xml[starts[row_idx]:ends[row_idx]]

        if isinstance(value, tuple):
            img_paths, caption = value   # (list[Path], str)
            cs, ce = _find_cell_boundaries(row_xml)
            cell_xml = row_xml[cs[col_idx]:ce[col_idx]]
            tc_pr_end = cell_xml.find("</w:tcPr>")
            prefix = (
                cell_xml[:tc_pr_end + len("</w:tcPr>")] if tc_pr_end != -1
                else cell_xml[:cell_xml.index(">") + 1]
            )
            inner = ""
            for j, img_path in enumerate(img_paths):
                if img_path not in img_cache:
                    safe_tag = re.sub(r"[^a-zA-Z0-9_-]", "_", f"custom_c{col_idx}_{img_path.stem}")
                    img_cache[img_path] = _register_image(img_path, tmp_dir, safe_tag)
                cx, cy = _image_size_emu(img_path, img_max_cm)
                drawing = _build_drawing(img_cache[img_path], cx, cy, 4000 + i * 100 + j, img_path.stem)
                inner += f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>{drawing}</w:drawing></w:r></w:p>'
            if caption:
                safe = caption.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                inner += f'<w:p><w:pPr><w:jc w:val="center"/></w:pPr>{_HIGHLIGHT_RUN.format(safe)}</w:p>'
            new_cell = prefix + inner + "</w:tc>"
            new_row = row_xml[:cs[col_idx]] + new_cell + row_xml[ce[col_idx]:]
        else:
            new_row = _set_text_in_row(row_xml, col_idx, str(value))

        patches.append((starts[row_idx], ends[row_idx], new_row))

    for s, e, new in reversed(patches):
        table_xml = table_xml[:s] + new + table_xml[e:]
    return table_xml


# =========================================================================== #
# ROW EXPANSION                                                                #
# =========================================================================== #

def _clone_empty_row(row_xml: str) -> str:
    """Return a copy of a data row with all cell text cleared."""
    return re.sub(r"<w:t(?:\s[^>]*)?>.*?</w:t>", "<w:t/>", row_xml, flags=re.DOTALL)


def _ensure_enough_rows(
    table_xml: str,
    needed: int,
    first_data: int,
    starts: list[int],
    ends: list[int],
) -> str:
    """Append cloned empty rows until the table has at least *needed* data rows."""
    available = len(starts) - first_data
    if available >= needed:
        return table_xml
    to_add = needed - available
    template_row = _clone_empty_row(table_xml[starts[-1]:ends[-1]])
    insert_pos = table_xml.rfind("</w:tbl>")
    if insert_pos == -1:
        insert_pos = len(table_xml)
    table_xml = table_xml[:insert_pos] + (template_row * to_add) + table_xml[insert_pos:]
    print(f"  Added {to_add} row(s) (had {available}, need {needed}).")
    return table_xml


# =========================================================================== #
# TABLE ORCHESTRATOR                                                            #
# =========================================================================== #

def fill_table(
    table_xml: str,
    component_names: list[str],
    template: dict[str, list[str]],
    tmp_dir: Path,
) -> str:
    """
    Fill the table using data from *template* (doc_mod_listado_componentes.json).
    Col 0 always gets the component names.
    Slots [0]..[N_DATA_COLS-1] map to table columns [1]..[N_DATA_COLS].
    String values are written as text; file-path values are embedded as images.
    """
    first_data, starts, ends = _detect_data_rows(table_xml)
    table_xml  = _ensure_enough_rows(table_xml, len(component_names), first_data, starts, ends)
    first_data, starts, ends = _detect_data_rows(table_xml)
    header_row = table_xml[starts[first_data - 1]:ends[first_data - 1]]

    cs, ce = _find_cell_boundaries(header_row)
    headers = [
        "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", header_row[s:e]))
        for s, e in zip(cs, ce)
    ]
    print(f"  {len(starts)} rows total, data starts at index {first_data}.")
    print(f"  Detected column headers:")
    for i, h in enumerate(headers):
        print(f"    [{i}] {h!r}")

    # Col 0: component names
    table_xml = _fill_first_column(table_xml, component_names, first_data, starts, ends)

    # Slots 0..N_DATA_COLS-1 → table columns 1..N_DATA_COLS
    # Re-detect after every modification because XML edits shift byte offsets.
    for slot in range(N_DATA_COLS):
        col_idx = slot + 1
        slot_data: dict[str, str | Path] = {}
        for name in component_names:
            slots = template.get(name, [])
            if slot < len(slots) and slots[slot].strip():
                v = slots[slot].strip()
                paths_str, _, cap = v.partition("||")
                img_paths = [Path(s.strip()) for s in paths_str.split(";;")]
                valid = [p for p in img_paths if p.is_file()]
                if valid:
                    slot_data[name] = (valid, cap)   # (list[Path], caption)
                else:
                    slot_data[name] = v
            else:
                slot_data[name] = "---"

        col_label = headers[col_idx] if col_idx < len(headers) else f"col {col_idx}"
        filled = sum(1 for v in slot_data.values() if v != "---")
        print(f"  Slot [{slot}] → col {col_idx} ({col_label!r}): {filled} values, {len(slot_data)-filled} empty ('---').")

        img_max_cm = MODELO_SIZE_CM if slot in (0, 1) else LOGO_SIZE_CM
        first_data, starts, ends = _detect_data_rows(table_xml)
        table_xml = apply_fills(
            table_xml, component_names, first_data, starts, ends, col_idx, slot_data, tmp_dir,
            img_max_cm=img_max_cm,
        )

    return table_xml


# =========================================================================== #
# MAIN                                                                         #
# =========================================================================== #

def _find_single_docx(folder: Path) -> Path:
    matches = [p for p in folder.glob("*.docx") if not p.name.startswith("~")]
    if not matches:
        raise FileNotFoundError(f"No .docx found in '{folder}'.")
    if len(matches) > 1:
        raise RuntimeError(f"Multiple .docx files in '{folder}': {[p.name for p in matches]}")
    return matches[0]


def main() -> None:
    template_path = JSON_FOLDER / "doc_mod_listado_componentes.json"
    if not template_path.is_file():
        print(f"Error: '{template_path}' not found.")
        print("Run tabla_LC_nombre.py first to generate doc_mod_listado_componentes.json.")
        raise SystemExit(1)

    template: dict[str, list[str]] = json.loads(template_path.read_text(encoding="utf-8"))
    component_names = list(template.keys())
    print(f"Components loaded from template: {len(component_names)}")

    informe_json = JSON_FOLDER / "doc_mod_numero_informe.json"
    if not informe_json.is_file():
        print(f"Error: '{informe_json}' not found.")
        raise SystemExit(1)
    docx_filename = json.loads(informe_json.read_text(encoding="utf-8")).get("filename")
    if not docx_filename:
        print(f"Error: 'filename' key missing in '{informe_json}'.")
        raise SystemExit(1)
    input_docx = DOCX_FOLDER / docx_filename
    if not input_docx.is_file():
        print(f"Error: '{input_docx}' not found.")
        raise SystemExit(1)

    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    output_docx = OUTPUT_FOLDER / input_docx.name

    print(f"\nInput : {input_docx}")
    print(f"Output: {output_docx}\n")

    tmp_dir = Path(tempfile.mkdtemp())
    try:
        with zipfile.ZipFile(input_docx, "r") as zin:
            zin.extractall(tmp_dir)

        doc_xml_path = tmp_dir / "word" / "document.xml"
        xml_content  = doc_xml_path.read_text(encoding="utf-8")

        tbl_start, tbl_end = _find_table_by_content(xml_content, "objeto")
        print("Table 4.1.2 found.")

        table_xml   = xml_content[tbl_start:tbl_end]
        table_xml   = fill_table(table_xml, component_names, template, tmp_dir)
        xml_content = _ensure_drawing_namespaces(
            xml_content[:tbl_start] + table_xml + xml_content[tbl_end:]
        )

        doc_xml_path.write_text(xml_content, encoding="utf-8")

        with zipfile.ZipFile(output_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for fp in tmp_dir.rglob("*"):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(tmp_dir).as_posix())

        print(f"\nDone -> {output_docx}")

        lista_folder = OUTPUT_FOLDER / ".docx"
        lista_folder.mkdir(parents=True, exist_ok=True)
        lista_docx = lista_folder / "doc_mod_taba_LC.docx"
        doc_xml_path.write_text(_build_table_only_xml(xml_content, table_xml), encoding="utf-8")
        with zipfile.ZipFile(lista_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for fp in tmp_dir.rglob("*"):
                if fp.is_file():
                    zout.write(fp, fp.relative_to(tmp_dir).as_posix())
        print(f"Done -> {lista_docx}")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
