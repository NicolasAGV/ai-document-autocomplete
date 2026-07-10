from pathlib import Path
import doc_mod_01_creation


input_folder = doc_mod_01_creation.output_folder
images_folder = doc_mod_01_creation.img_renamed_folder
output_folder = doc_mod_01_creation.output_folder

# Fallback template used when documento_informe.json (or the docx it points to) is missing.
FALLBACK_DOCX = Path(__file__).parent / ".XLSX .DOCX patrones" / "Generico - Anexo 5 Fotos.docx"

"""
modify_docx_images.py  —  Part 6
==================================
Inserts renamed images (from extract_images.py output) into the .docx:

  1. All "Marcado_Producto*.jpg"   → "Reproducción de la(s) placa(s) de identificación" cell
  2. "Certificadora_id.jpg"   → "Reproducción de la etiqueta de la certificadora" cell
  3. All remaining images          → ANEXO 5: FOTOS grid (2 cols × 3 rows per page)
     - Order: General, Marcado_Producto, Ficha_de_alimentacion, Cordon_de_alimentacion,
              Conector_de_alimentacion, Zocalo_de_alimentacion, Selector_de_alimentacion,
              Interruptor, PCB, Fusible, Varistor, Capacitor_X, Bobina,
              Resistencia, Capacitor_Electrolitico, Capacitor_Y, Optoacoplador,
            Transformador, Ventilador, Parlante, varios
     - Logo images are excluded from the grid
     - "varios" images are skipped (already excluded by extract_images.py)
     - Last page: only as many cells as images (no blank cells left over)
     - Final-de-documento image preserved at the very end

Grid structure (from Example_matrix_images.docx):
  PIC row  : h=3969 DXA (exact), cols = 4253 | 284 | 4253
  SPACE row: h=284  DXA,         cols = 4253 | 284 | 4253

Usage:
    python modify_docx_images.py  <docx>  <fotos_renamed_folder>  [output.docx]

    <docx>                  the original (or already-modified) .docx
    <fotos_renamed_folder>  folder produced by extract_images.py  (fotos_renamed/)
                            OR a Fotos_Renamed.zip — either works
    [output.docx]           optional; defaults to <docx_stem>_images.docx

Install dependencies (once):
    pip install lxml Pillow

Import as module:
    from modify_docx_images import insert_images
    insert_images("report.docx", "fotos_renamed/", "report_images.docx")
"""

import re
import sys
import shutil
import zipfile
import tempfile
# from pathlib import Path

try:
    from lxml import etree
except ImportError:
    print("ERROR: lxml not installed. Run:  pip install lxml"); sys.exit(1)
try:
    from PIL import Image as PILImage
except ImportError:
    print("ERROR: Pillow not installed. Run:  pip install Pillow"); sys.exit(1)


# ── Namespace constants ───────────────────────────────────────────────────────
W       = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP      = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A       = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
PIC_NS  = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
R       = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
W14     = '{http://schemas.microsoft.com/office/word/2010/wordml}'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
REL_IMG = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
VML_NS  = '{urn:schemas-microsoft-com:vml}'

# ── Grid layout constants (from Example_matrix_images.docx) ──────────────────
COL_PIC  = 4253   # DXA — wide image column
COL_SEP  = 284    # DXA — narrow separator column
ROW_PIC  = 3969   # DXA — tall image row (exact height)
ROW_SEP  = 284    # DXA — short separator row

# Max image size in EMU (1 DXA = 914400/1440 EMU)
MAX_IMG_W = int(COL_PIC * 914400 / 1440)
MAX_IMG_H = int(ROW_PIC * 914400 / 1440)

# Placa/cert cell max sizes in EMU
PLACA_MAX_W = int(9889 * 914400 / 1440)
PLACA_MAX_H = int(4768 * 914400 / 1440)
CERT_MAX_W  = int(9883 * 914400 / 1440)
CERT_MAX_H  = int(2473 * 914400 / 1440)

# ── User settings ────────────────────────────────────────────────────────────
# Set True to completely skip Marcado_Producto and Certificadora_id images
# (they will not be inserted into the placa/cert cells nor into the grid).
IGNORE_MARCADO_IMAGES = False

# Grid image order (prefix of renamed filename)
GRID_ORDER = [
    'General', 'Marcado_Producto', 'Ficha_de_alimentacion', 'Cordon_de_alimentacion',
    'Conector_de_alimentacion', 'Zocalo_de_alimentacion', 'Selector_de_alimentacion',
    'Interruptor', 'PCB', 'Fusible', 'Varistor', 'Capacitor_X', 'Bobina',
    'Resistencia', 'Capacitor_Electrolitico', 'Capacitor_Y', 'Optoacoplador',
    'Transformador', 'Ventilador', 'Parlante', 'varios',
]


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers — counters (mutable lists so they work as shared state)
# ═══════════════════════════════════════════════════════════════════════════════

def _next(counter: list) -> int:
    v = counter[0]; counter[0] += 1; return v


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers — relationships
# ═══════════════════════════════════════════════════════════════════════════════

def add_image_rel(rels_root: etree.Element, media_dir: Path,
                  img_path: Path, rid_counter: list,
                  img_counter: list) -> str:
    """Copy image to media/, add a relationship entry, return the rId string."""
    ext   = img_path.suffix.lower()
    mname = f'image{_next(img_counter)}{ext}'
    shutil.copy2(img_path, media_dir / mname)

    rid = f'rId{_next(rid_counter)}'
    el  = etree.SubElement(rels_root, f'{{{RELS_NS}}}Relationship')
    el.set('Id',     rid)
    el.set('Type',   REL_IMG)
    el.set('Target', f'media/{mname}')
    return rid


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers — image sizing
# ═══════════════════════════════════════════════════════════════════════════════

def fit_emu(img_path: Path, max_w: int, max_h: int) -> tuple:
    """Return (cx, cy) in EMU, scaled to fit within max_w × max_h."""
    with PILImage.open(img_path) as im:
        w_px, h_px = im.size
    w_emu = w_px * 9525   # 9525 EMU per pixel at 96 dpi
    h_emu = h_px * 9525
    scale = min(max_w / w_emu, max_h / h_emu, 1.0)
    return int(w_emu * scale), int(h_emu * scale)


# ═══════════════════════════════════════════════════════════════════════════════
# Helpers — XML building blocks
# ═══════════════════════════════════════════════════════════════════════════════

def make_drawing(rid: str, cx: int, cy: int,
                 draw_counter: list) -> etree.Element:
    """Build a <w:drawing><wp:inline> element embedding an image."""
    d = _next(draw_counter)
    xml = (
        f'<w:drawing'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
        f' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
        f' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
        f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{d}" name="Img{d}"/>'
        f'<wp:cNvGraphicFramePr>'
        f'<a:graphicFrameLocks noChangeAspect="1"/>'
        f'</wp:cNvGraphicFramePr>'
        f'<a:graphic>'
        f'<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic>'
        f'<pic:nvPicPr>'
        f'<pic:cNvPr id="{d}" name="Img{d}"/>'
        f'<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr>'
        f'</pic:nvPicPr>'
        f'<pic:blipFill>'
        f'<a:blip r:embed="{rid}"/>'
        f'<a:stretch><a:fillRect/></a:stretch>'
        f'</pic:blipFill>'
        f'<pic:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'</pic:spPr>'
        f'</pic:pic>'
        f'</a:graphicData>'
        f'</a:graphic>'
        f'</wp:inline>'
        f'</w:drawing>'
    )
    return etree.fromstring(xml)


def make_img_paragraph(rid: str, img_path: Path,
                       max_w: int, max_h: int,
                       draw_counter: list) -> etree.Element:
    """Build a centered paragraph containing a single image drawing."""
    cx, cy = fit_emu(img_path, max_w, max_h)
    p = etree.fromstring(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:pPr><w:jc w:val="center"/></w:pPr>'
        '<w:r/>'
        '</w:p>')
    p.find(f'{W}r').append(make_drawing(rid, cx, cy, draw_counter))
    return p


def _pid(pid_counter: list) -> str:
    return f'{_next(pid_counter):08X}'


# ═══════════════════════════════════════════════════════════════════════════════
# Grid table builders
# ═══════════════════════════════════════════════════════════════════════════════

def _pic_cell(img: Path | None,
              rels_root, media_dir,
              rid_counter, img_counter, draw_counter,
              pid_counter) -> etree.Element:
    """Build a PIC table cell, with or without an image."""
    tc = etree.fromstring(
        f'<w:tc'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        f'<w:tcPr>'
        f'<w:tcW w:w="{COL_PIC}" w:type="dxa"/>'
        f'<w:vAlign w:val="center"/>'
        f'</w:tcPr>'
        f'<w:p w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'</w:p>'
        f'</w:tc>')
    if img is not None:
        rid = add_image_rel(rels_root, media_dir, img,
                            rid_counter, img_counter)
        cx, cy = fit_emu(img, MAX_IMG_W, MAX_IMG_H)
        p = tc.find(f'{W}p')
        r = etree.SubElement(p, f'{W}r')
        r.append(make_drawing(rid, cx, cy, draw_counter))
    return tc


def _pic_cell_spanning(img: Path,
                       rels_root, media_dir,
                       rid_counter, img_counter, draw_counter,
                       pid_counter) -> etree.Element:
    """Build a PIC cell spanning all 3 grid columns (lone last image, no empty cell)."""
    full_w = COL_PIC + COL_SEP + COL_PIC
    tc = etree.fromstring(
        f'<w:tc'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        f'<w:tcPr>'
        f'<w:tcW w:w="{full_w}" w:type="dxa"/>'
        f'<w:gridSpan w:val="3"/>'
        f'<w:vAlign w:val="center"/>'
        f'</w:tcPr>'
        f'<w:p w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'</w:p>'
        f'</w:tc>')
    rid = add_image_rel(rels_root, media_dir, img, rid_counter, img_counter)
    cx, cy = fit_emu(img, MAX_IMG_W, MAX_IMG_H)
    p = tc.find(f'{W}p')
    r = etree.SubElement(p, f'{W}r')
    r.append(make_drawing(rid, cx, cy, draw_counter))
    return tc


def _sep_cell(pid_counter) -> etree.Element:
    """Build the narrow SPACE separator cell."""
    return etree.fromstring(
        f'<w:tc'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
        f'<w:tcPr>'
        f'<w:tcW w:w="{COL_SEP}" w:type="dxa"/>'
        f'<w:vAlign w:val="center"/>'
        f'</w:tcPr>'
        f'<w:p w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:pPr><w:ind w:right="-81"/><w:jc w:val="center"/></w:pPr>'
        f'</w:p>'
        f'</w:tc>')


def _pic_row(left: Path | None, right: Path | None,
             rels_root, media_dir,
             rid_counter, img_counter, draw_counter,
             pid_counter) -> etree.Element:
    """Build a PIC row (tall, exact height) with left and optional right image."""
    tr = etree.fromstring(
        f'<w:tr'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        f' w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:trPr>'
        f'<w:trHeight w:hRule="exact" w:val="{ROW_PIC}"/>'
        f'<w:jc w:val="center"/>'
        f'</w:trPr>'
        f'</w:tr>')
    if right is None:
        tr.append(_pic_cell_spanning(left, rels_root, media_dir,
                                     rid_counter, img_counter, draw_counter, pid_counter))
    else:
        tr.append(_pic_cell(left,  rels_root, media_dir,
                            rid_counter, img_counter, draw_counter, pid_counter))
        tr.append(_sep_cell(pid_counter))
        tr.append(_pic_cell(right, rels_root, media_dir,
                            rid_counter, img_counter, draw_counter, pid_counter))
    return tr


def _space_row(pid_counter) -> etree.Element:
    """Build a short SPACE separator row between PIC rows."""
    tr = etree.fromstring(
        f'<w:tr'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        f' w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:trPr>'
        f'<w:trHeight w:val="{ROW_SEP}"/>'
        f'<w:jc w:val="center"/>'
        f'</w:trPr>'
        f'</w:tr>')
    for col_w in [COL_PIC, COL_SEP, COL_PIC]:
        tc = etree.fromstring(
            f'<w:tc'
            f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">'
            f'<w:tcPr>'
            f'<w:tcW w:w="{col_w}" w:type="dxa"/>'
            f'<w:vAlign w:val="center"/>'
            f'</w:tcPr>'
            f'<w:p w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
            f'<w:pPr><w:ind w:right="-81"/><w:jc w:val="center"/></w:pPr>'
            f'</w:p>'
            f'</w:tc>')
        tr.append(tc)
    return tr


def build_photo_table(pairs: list,
                      rels_root, media_dir,
                      rid_counter, img_counter,
                      draw_counter, pid_counter) -> etree.Element:
    """
    Build one full photo table for a page.
    pairs = list of (left_img, right_img) where right_img may be None.
    """
    tbl = etree.fromstring(
        f'<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:tblPr>'
        f'<w:tblW w:w="0" w:type="auto"/>'
        f'<w:jc w:val="center"/>'
        f'<w:tblLayout w:type="fixed"/>'
        f'<w:tblCellMar>'
        f'<w:left w:w="0" w:type="dxa"/>'
        f'<w:right w:w="0" w:type="dxa"/>'
        f'</w:tblCellMar>'
        f'<w:tblLook w:val="0000"'
        f' w:firstRow="0" w:lastRow="0"'
        f' w:firstColumn="0" w:lastColumn="0"'
        f' w:noHBand="0" w:noVBand="0"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>'
        f'<w:gridCol w:w="{COL_PIC}"/>'
        f'<w:gridCol w:w="{COL_SEP}"/>'
        f'<w:gridCol w:w="{COL_PIC}"/>'
        f'</w:tblGrid>'
        f'</w:tbl>')

    for i, (left, right) in enumerate(pairs):
        tbl.append(_pic_row(left, right,
                            rels_root, media_dir,
                            rid_counter, img_counter,
                            draw_counter, pid_counter))
        if i < len(pairs) - 1:
            tbl.append(_space_row(pid_counter))

    return tbl


def page_break_para(pid_counter) -> etree.Element:
    return etree.fromstring(
        f'<w:p'
        f' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
        f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
        f' w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
        f'<w:pPr><w:jc w:val="center"/></w:pPr>'
        f'<w:r><w:br w:type="page"/></w:r>'
        f'</w:p>')


# ═══════════════════════════════════════════════════════════════════════════════
# Image sorting
# ═══════════════════════════════════════════════════════════════════════════════

def grid_sort_key(path: Path) -> tuple:
    """Sort grid images by GRID_ORDER, then by number, then General before Marcado."""
    stem = path.stem
    for i, group in enumerate(GRID_ORDER):
        if stem.startswith(group):
            m = re.search(r'_(\d+)', stem)
            num = int(m.group(1)) if m else 0
            suf = 0 if 'General' in stem else 1
            return (i, num, suf)
    return (99, 0, 0)


def collect_images(img_folder: Path) -> dict:
    """
    Scan img_folder and return categorised image paths:
        {
          'placa':  [Path, ...],   # Marcado_Producto*.jpg
          'cert':   Path | None,   # Certificadora_id.jpg
          'grid':   [Path, ...],   # everything else (sorted, no Logo, no varios)
        }
    """
    IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}

    # Use a dict keyed on lowercase filename to avoid duplicates on
    # case-insensitive file systems (Windows glob returns *.jpg AND *.JPG
    # for the same file, causing every image to be inserted twice).
    seen = {}
    for f in img_folder.iterdir():
        if f.is_file() and f.suffix.lower() in IMAGE_EXTS:
            seen[f.name.lower()] = f
    all_imgs = sorted(seen.values(), key=lambda p: p.name.lower())

    placa  = []
    cert   = None
    grid   = []

    for p in all_imgs:
        stem = p.stem
        if 'Marcado_Producto' in stem:
            placa.append(p)
        elif 'certificadora_id' in stem.lower():
            cert = p
        elif 'Logo' in stem or 'varios' in stem.lower():
            continue   # skip
        else:
            grid.append(p)

    grid.sort(key=grid_sort_key)
    return {'placa': placa, 'cert': cert, 'grid': grid}


# ═══════════════════════════════════════════════════════════════════════════════
# Docx unpack / pack
# ═══════════════════════════════════════════════════════════════════════════════

def unpack(docx: Path, dest: Path):
    dest.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx) as zf:
        zf.extractall(dest)


def pack(unpacked: Path, out_docx: Path):
    if out_docx.exists(): out_docx.unlink()
    with zipfile.ZipFile(out_docx, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(unpacked.rglob('*')):
            if f.is_file():
                zf.write(f, f.relative_to(unpacked))


# ═══════════════════════════════════════════════════════════════════════════════
# Main function
# ═══════════════════════════════════════════════════════════════════════════════

def insert_images(docx_path: str | Path,
                  imgs_source: str | Path,
                  output_path: str | Path | None = None) -> Path:
    """
    Insert images into docx_path and write the result to output_path.

    imgs_source can be:
        - a folder produced by extract_images.py  (fotos_renamed/)
        - a Fotos_Renamed.zip file

    Returns the path to the output docx.
    """
    docx_path = Path(docx_path)
    imgs_src  = Path(imgs_source)

    if output_path is None:
        output_path = docx_path
    output_path = Path(output_path)

    # ── If imgs_source is a zip, extract it first ─────────────────────────────
    tmp_imgs = None
    if imgs_src.suffix.lower() == '.zip':
        tmp_imgs = Path(tempfile.mkdtemp()) / 'fotos_renamed'
        tmp_imgs.mkdir()
        with zipfile.ZipFile(imgs_src) as zf:
            zf.extractall(tmp_imgs)
        img_folder = tmp_imgs
    else:
        img_folder = imgs_src

    # ── Categorise images ─────────────────────────────────────────────────────
    images = collect_images(img_folder)
    if IGNORE_MARCADO_IMAGES:
        images['placa'] = []
        images['cert']  = None
    print(f"  Placa images : {len(images['placa'])}")
    print(f"  Cert image   : {images['cert'].name if images['cert'] else 'NOT FOUND'}")
    print(f"  Grid images  : {len(images['grid'])}")
    for g in images['grid']: print(f"    {g.name}")

    # ── Unpack docx ───────────────────────────────────────────────────────────
    tmp_docx = Path(tempfile.mkdtemp()) / 'unpacked'
    unpack(docx_path, tmp_docx)

    doc_path  = tmp_docx / 'word' / 'document.xml'
    rels_path = tmp_docx / 'word' / '_rels' / 'document.xml.rels'
    media_dir = tmp_docx / 'word' / 'media'
    media_dir.mkdir(exist_ok=True)

    doc_tree  = etree.parse(str(doc_path))
    root_el   = doc_tree.getroot()
    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()

    # ── Initialise counters ───────────────────────────────────────────────────
    all_rids = [int(re.search(r'\d+', el.get('Id')).group())
                for el in rels_root if re.search(r'\d+', el.get('Id', ''))]
    all_imgn = [int(re.search(r'\d+', p.stem).group())
                for p in media_dir.iterdir() if re.search(r'\d+', p.stem)]

    rid_counter  = [max(all_rids, default=44) + 1]
    img_counter  = [max(all_imgn, default=32) + 1]
    draw_counter = [500]
    pid_counter  = [0x70000001]

    # ── 1. Insert Marcado_Producto into placa cell ────────────────────────────
    PLACA_LABEL = 'Reproducción de la(s) placa(s) de identificación'
    cell_placa  = None
    for tc in root_el.iter(f'{W}tc'):
        if any(PLACA_LABEL in (t.text or '') for t in tc.iter(f'{W}t')):
            cell_placa = tc; break

    if cell_placa is not None and images['placa']:
        # Remove empty paragraphs (keep the one with the label text)
        for p in list(cell_placa.findall(f'{W}p')):
            if not any(t.text and t.text.strip() for t in p.iter(f'{W}t')):
                cell_placa.remove(p)
        # Divide width equally between all placa images
        per_w = PLACA_MAX_W // len(images['placa'])
        for img in images['placa']:
            rid = add_image_rel(rels_root, media_dir, img,
                                rid_counter, img_counter)
            cell_placa.append(
                make_img_paragraph(rid, img, per_w, PLACA_MAX_H, draw_counter))
        print(f"\n  ✔ Placa cell: inserted {len(images['placa'])} image(s)")
    else:
        if cell_placa is None:
            print("\n  ⚠ Placa cell not found")
        if not images['placa']:
            print("\n  ⚠ No Marcado_Producto images found")

    # ── 2. Insert Certificadora_id into cert cell ────────────────────────
    CERT_LABEL = 'Reproducción de la etiqueta de la certificadora'
    cell_cert  = None
    for tc in root_el.iter(f'{W}tc'):
        if any(CERT_LABEL in (t.text or '') for t in tc.iter(f'{W}t')):
            cell_cert = tc; break

    if cell_cert is not None and images['cert'] is not None:
        for p in list(cell_cert.findall(f'{W}p')):
            if not any(t.text and t.text.strip() for t in p.iter(f'{W}t')):
                cell_cert.remove(p)
        rid = add_image_rel(rels_root, media_dir, images['cert'],
                            rid_counter, img_counter)
        cell_cert.append(
            make_img_paragraph(rid, images['cert'],
                               CERT_MAX_W, CERT_MAX_H, draw_counter))
        print(f"  ✔ Cert cell:  inserted {images['cert'].name}")
    else:
        if cell_cert is None:
            print("  ⚠ Cert cell not found")
        if images['cert'] is None:
            print("  ⚠ Certificadora_id image not found")

    # ── 3. Build ANEXO 5 grid ─────────────────────────────────────────────────
    grid = images['grid']
    body = root_el.find(f'{W}body')
    body_children = list(body)

    # Find ANEXO 5 heading paragraph
    # Use loose matching: Word may store the number as a field code or split
    # it across runs, so 'ANEXO 5' may not appear as a single joined string.
    anexo5_idx = None
    for i, ch in enumerate(body_children):
        if ch.tag == f'{W}p':
            txt = ''.join(t.text or '' for t in ch.iter(f'{W}t'))
            if re.search(r'ANEXO\s*5', txt) and 'FOTO' in txt:
                anexo5_idx = i; break
    if anexo5_idx is None:
        # Fallback: accept any paragraph that contains both ANEXO and FOTO
        for i, ch in enumerate(body_children):
            if ch.tag == f'{W}p':
                txt = ''.join(t.text or '' for t in ch.iter(f'{W}t'))
                if 'ANEXO' in txt and 'FOTO' in txt:
                    anexo5_idx = i; break

    if anexo5_idx is None:
        print("  ⚠ ANEXO 5 heading not found — skipping grid")
        print("  DEBUG: all non-empty paragraph texts anywhere in document:")
        for p in root_el.iter(f'{W}p'):
            txt = ''.join(t.text or '' for t in p.iter(f'{W}t'))
            if txt.strip():
                print(f"    {repr(txt)}")
    else:
        # Save final-de-documento paragraph (has a VML imagedata element)
        final_para = None
        for ch in reversed(list(body)):
            if ch.tag == f'{W}p' and \
               ch.find(f'.//{VML_NS}imagedata') is not None:
                final_para = ch; break

        # Determine the end boundary in the snapshot (before final_para / sectPr)
        sectPr = body.find(f'{W}sectPr')
        end_idx = len(body_children)
        if final_para is not None and final_para in body_children:
            end_idx = min(end_idx, body_children.index(final_para))
        if sectPr is not None and sectPr in body_children:
            end_idx = min(end_idx, body_children.index(sectPr))

        # Remove ALL placeholder photo tables between the heading and the boundary
        # (template may have one table per page; we only keep what we rebuild)
        old_tables = [node for node in body_children[anexo5_idx + 1 : end_idx]
                      if node.tag == f'{W}tbl']

        # Temporarily remove final_para and sectPr so we can insert before them
        if final_para is not None: body.remove(final_para)
        if sectPr     is not None: body.remove(sectPr)
        for tbl in old_tables:
            body.remove(tbl)

        # Split grid images into pages of 6 (3 rows × 2 cols)
        # Last page: only fill as many cells as images available (no blanks)
        pages = []
        for i in range(0, len(grid), 6):
            chunk = grid[i:i + 6]
            pairs = [
                (chunk[j], chunk[j + 1] if j + 1 < len(chunk) else None)
                for j in range(0, len(chunk), 2)
            ]
            pages.append(pairs)

        # Insert tables into body
        insert_pos = list(body).index(body_children[anexo5_idx]) + 1
        for pi, pairs in enumerate(pages):
            if pi > 0:
                body.insert(insert_pos, page_break_para(pid_counter))
                insert_pos += 1
            tbl = build_photo_table(pairs,
                                    rels_root, media_dir,
                                    rid_counter, img_counter,
                                    draw_counter, pid_counter)
            body.insert(insert_pos, tbl)
            insert_pos += 1

        # Re-attach final_para with one blank line above it, then sectPr.
        # Strip pageBreakBefore and any large spacing so it follows immediately
        # after the last image rather than being pushed to a new page.
        if final_para is not None:
            pPr = final_para.find(f'{W}pPr')
            if pPr is None:
                pPr = etree.SubElement(final_para, f'{W}pPr')
                final_para.insert(0, pPr)
            pb = pPr.find(f'{W}pageBreakBefore')
            if pb is not None:
                pPr.remove(pb)
            sp = pPr.find(f'{W}spacing')
            if sp is not None:
                sp.attrib.pop(f'{W}before', None)
                sp.attrib.pop(f'{W}beforeLines', None)
            body.append(etree.fromstring(
                f'<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                f' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
                f' w14:paraId="{_pid(pid_counter)}" w14:textId="77777777">'
                f'<w:pPr><w:jc w:val="center"/>'
                f'<w:spacing w:before="0" w:after="0"/>'
                f'</w:pPr>'
                f'</w:p>'))
            body.append(final_para)
        if sectPr is not None: body.append(sectPr)

        n_pages = len(pages)
        n_imgs  = len(grid)
        print(f"  ✔ Grid: {n_imgs} images across {n_pages} page(s)")

    # ── Fix Content-Types for .jpg if not already present ─────────────────────
    ct_path = tmp_docx / '[Content_Types].xml'
    ct_text = ct_path.read_text()
    if 'Extension="jpg"' not in ct_text:
        ct_text = ct_text.replace(
            '</Types>',
            '<Default Extension="jpg" ContentType="image/jpeg"/></Types>')
        ct_path.write_text(ct_text)

    # ── Save XML ──────────────────────────────────────────────────────────────
    doc_tree.write(str(doc_path),
                   xml_declaration=True, encoding='UTF-8', standalone=True)
    rels_tree.write(str(rels_path),
                    xml_declaration=True, encoding='UTF-8', standalone=True)

    # ── Pack docx ─────────────────────────────────────────────────────────────
    pack(tmp_docx, output_path)
    print(f"\n  ✔ Output: {output_path}")

    # ── Cleanup ───────────────────────────────────────────────────────────────
    shutil.rmtree(tmp_docx)
    if tmp_imgs is not None:
        shutil.rmtree(tmp_imgs.parent)

    return output_path


def _save_image_tables_docx(source_docx: Path, dest_docx: Path) -> None:
    """Save a docx containing only the tables that have embedded images."""
    tmp = Path(tempfile.mkdtemp()) / 'unpacked'
    unpack(source_docx, tmp)
    doc_path = tmp / 'word' / 'document.xml'
    doc_tree = etree.parse(str(doc_path))
    root_el  = doc_tree.getroot()
    body     = root_el.find(f'{W}body')

    keep    = [ch for ch in list(body)
               if ch.tag == f'{W}tbl' and ch.find(f'.//{WP}inline') is not None]
    sect_pr = body.find(f'{W}sectPr')

    for child in list(body):
        body.remove(child)
    for tbl in keep:
        body.append(tbl)
    if sect_pr is not None:
        body.append(sect_pr)

    doc_tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)
    pack(tmp, dest_docx)
    shutil.rmtree(tmp)
    print(f"  ✔ Image-tables copy: {dest_docx}")


# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder where the input .docx is located
    # input_folder  = Path(r'C:\Users\yourname\Documents\job_folder')

    # # Folder where fotos_renamed/ is located (produced by extract_images.py)
    # images_folder = input_folder / 'output' / 'fotos_renamed'

    # # Folder where the output .docx will be written
    # output_folder = input_folder / 'output'
    # ── End of configuration ──────────────────────────────────────────────────

    output_folder.mkdir(parents=True, exist_ok=True)

    # Find input .docx from doc_mod_numero_informe.json
    import json as _json
    informe_json = doc_mod_01_creation.json_folder / "doc_mod_numero_informe.json"
    docx_file = None
    if informe_json.exists():
        with open(informe_json, encoding='utf-8') as f:
            _inf = _json.load(f)
        docx_filename = _inf.get("filename")
        if docx_filename:
            candidate = input_folder / docx_filename
            if candidate.exists():
                docx_file = candidate
            else:
                print(f"Warning: '{candidate}' not found. Falling back to generic template.")
        else:
            print(f"Warning: 'filename' key missing in '{informe_json}'. Falling back to generic template.")
    else:
        print(f"Warning: '{informe_json}' not found. Falling back to generic template.")
    if docx_file is None:
        docx_file = FALLBACK_DOCX
        if not docx_file.exists():
            print(f"ERROR: Fallback '{docx_file}' also not found.")
            sys.exit(1)
    print(f"\n  Found docx: {docx_file.name}")

    if not images_folder.exists():
        print(f"ERROR: images folder not found: {images_folder}")
        sys.exit(1)

    print(f"  File   : {docx_file}")
    print(f"  Images : {images_folder}\n")

    insert_images(docx_file, images_folder)

    docx_subfolder = output_folder / ".docx"
    docx_subfolder.mkdir(parents=True, exist_ok=True)
    _save_image_tables_docx(docx_file, docx_subfolder / "doc_mod_images.docx")

    print("\nDone.\n")