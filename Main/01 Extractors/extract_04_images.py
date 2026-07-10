# EN MODO STAND ALONE: 

from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

images_folder = Path(_mp.main_path)
output_folder = images_folder / 'output'
json_path = output_folder / 'json' / 'extracted_listado_componentes_raw.json'


# =============================================================================
# NOMENCLATURE RULES  —  edit this section to update recognition logic
# =============================================================================
#
# EJEMPLOS RAPIDOS
# gnr6
# mar1
# cer
# cor1m3
# fus1l1
# opt2g2
#
#
# STANDARD FILENAME FORMATS
# -------------------------
# Two valid structures are recognised:
#
#   FORMAT A — multiple photos of the same type:
#   {3-letter code} {1 digit} {1 letter} {1 digit}
#        |              |          |          |
#   component type  component  photo type  photo
#     (see map)       id #     g/m/l        id #
#
#   Example:  cor1m2.jpg
#             cor → Cordon_de_alimentacion
#             1   → component identity #1
#             m   → Marcado
#             2   → photo identity #2
#             → output: Cordon_de_alimentacion_1_Marcado_2.jpg
#
#   FORMAT B — single photo of that type (no photo id needed):
#   {3-letter code} {1 digit} {1 letter}
#        |              |          |
#   component type  component  photo type
#     (see map)       id #     g/m/l
#
#   Example:  fic1g.jpg
#             fic → Ficha_de_alimentacion
#             1   → component identity #1
#             g   → General
#             → output: Ficha_de_alimentacion_1_General.jpg
#
#
# COMPONENT CODE MAP  (3-letter prefix → component name)
# -------------------------------------------------------
#   bob → Bobina
#   bor → Bornera
#   cae → Capacitor_Electrolitico   (alt: cel)
#   cax → Capacitor_X
#   cay → Capacitor_Y
#   con → Conector_de_alimentacion
#   cor → Cordon_de_alimentacion
#   fic → Ficha_de_alimentacion
#   fus → Fusible
#   gnr → General
#   int → Interruptor
#   opt → Optoacoplador
#   par → Parlante
#   pla → PCB
#   res → Resistencia
#   sel → Selector_de_alimentacion
#   tra → Transformador
#   var → Varistor
#   ven → Ventilador
#   cer → Certificadora_id
#   zoc → Zocalo_de_alimentacion    (alt: zol)
#
#
# PHOTO TYPE LETTERS
# ------------------
#   g → General
#   m → Marcado
#   l → Logo
#
#
# SPECIAL CASES  (prefix + optional photo number, no component id or type)
# -------------------------------------------------------------------------
#   These prefixes follow the pattern:  {prefix}{optional digit}
#
#   mar       → Marcado_Producto
#   mar1      → Marcado_Producto_1
#   mar2      → Marcado_Producto_2
#
#   gnr       → General
#   gnr1      → General_1
#   gnr2      → General_2
#
#   cer       → Certificadora_id   (exact only, no digit variant)
#
#
# UNRECOGNISED FILES
# ------------------
#   Any filename that does not match the standard format or the special
#   cases above is labelled:  Varios1.jpg, Varios2.jpg, ...
#
#   Files whose name contains the word "varios" are skipped entirely
#   (not copied to the output folder).
#
# =============================================================================

import json
import re
import sys
import shutil
import zipfile


# ── Nomenclature map (3-letter codes only) ───────────────────────────────────
COMPONENT_MAP = {
    'cer': 'Certificadora_id',
    'mar': 'Marcado_Producto',
    'cae': 'Capacitor_Electrolitico',
    'cel': 'Capacitor_Electrolitico',
    'cax': 'Capacitor_X',
    'cay': 'Capacitor_Y',
    'res': 'Resistencia',
    'bob': 'Bobina',
    'bor': 'Bornera',
    'fus': 'Fusible',
    'gnr': 'General',
    'opt': 'Optoacoplador',
    'pla': 'PCB',
    'tra': 'Transformador',
    'cor': 'Cordon_de_alimentacion',
    'fic': 'Ficha_de_alimentacion',
    'con': 'Conector_de_alimentacion',    
    'zoc': 'Zocalo_de_alimentacion',
    'sel': 'Selector_de_alimentacion',
    'int': 'Interruptor',
    'par': 'Parlante',
    'var': 'Varistor',
    'ven': 'Ventilador',
}

SUFFIX_MAP = {
    'g': 'General',
    'm': 'Marcado',
    'l': 'Logo',
}

IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif', '.webp'}


def rename_image(stem: str) -> str | None:
    """
    Given a file stem (lowercase, no extension), return the new base name
    without extension, or None if the file should be skipped (varios).
    Returns '' if unrecognised (caller will use "VariosN").

    Format A: {3-letter code}{component#}{g|m|l}{photo#}
    Format B: {3-letter code}{component#}{g|m|l}          (single photo, no id)
    Special:  mar{digit?} | gnr{digit?} | cer (exact)
    """
    if 'varios' in stem:
        return None

    s = stem.lower()

    if s == 'cer':
        return 'Certificadora_id'

    m_special = re.fullmatch(r'(mar|gnr)(\d?)', s)
    if m_special:
        prefix, num = m_special.groups()
        name = COMPONENT_MAP[prefix]
        return f'{name}_{num}' if num else name

    m = re.fullmatch(r'([a-z]{3})(\d)([gml])(\d?)', s)
    if m:
        code, comp_num, suffix_letter, pic_num = m.groups()
        component = COMPONENT_MAP.get(code)
        if component:
            base = f'{component}_{comp_num}_{SUFFIX_MAP[suffix_letter]}'
            return f'{base}_{pic_num}' if pic_num else base

    return ''   # unrecognised → Varios


def extract(source: str | Path,
            output_folder: str | Path) -> dict:
    """
    Rename images from source (a .zip file OR a folder) and write
    the renamed images as loose files into output_folder/fotos_renamed/.

    Returns a dict:  { new_filename: Path_to_renamed_file }
    """
    source        = Path(source)
    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True)

    out_dir = output_folder / 'fotos_renamed'
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir()

    # ── Collect raw images from zip or folder ─────────────────────────────────
    staging = None

    if source.is_file() and source.suffix.lower() == '.zip':
        # Extract zip to a temporary staging area
        staging = output_folder / '_staging'
        if staging.exists():
            shutil.rmtree(staging)
        with zipfile.ZipFile(source) as zf:
            zf.extractall(staging)
        search_root = staging
    elif source.is_dir():
        search_root = source
    else:
        print(f"  ERROR: source is neither a .zip nor a folder: {source}")
        return {}

    raw_images = sorted([
        f for f in search_root.rglob('*')
        if f.is_file() and f.suffix.lower() in IMAGE_EXTENSIONS
    ])

    # ── Rename and copy to out_dir ────────────────────────────────────────────
    renamed    = {}
    recognised = {}
    unmatched  = 0

    for src in raw_images:
        stem     = src.stem.lower()
        new_stem = rename_image(stem)

        if new_stem is None:
            print(f"  SKIP  (varios): {src.name}")
            continue

        if new_stem == '':
            unmatched += 1
            new_stem = f'Varios{unmatched}'
            print(f"  ????  {src.name:30s} → {new_stem}.jpg")
        else:
            print(f"  OK    {src.name:30s} → {new_stem}.jpg")

        new_name = new_stem + '.jpg'
        dst      = out_dir / new_name
        shutil.copy2(src, dst)
        renamed[new_name] = dst
        if not new_stem.startswith('Varios'):
            recognised[new_name] = dst

    # ── Cleanup staging if we created one ─────────────────────────────────────
    if staging is not None:
        shutil.rmtree(staging)

    skipped = len(raw_images) - len(renamed)
    print(f"\n  Renamed : {len(renamed)} images")
    print(f"  Skipped : {skipped} images")
    print(f"  Output  : {out_dir}")

    # json_path = output_folder / 'fotos_renamed.json'
    json_path.parent.mkdir(parents=True, exist_ok=True)
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(sorted(recognised.keys()), f, ensure_ascii=False, indent=2)
    print(f"  JSON    : {json_path}")

    return renamed

# ── CLI ───────────────────────────────────────────────────────────────────────
if __name__ == '__main__':

    # ── Configuration ─────────────────────────────────────────────────────────
    # Folder containing the raw images to rename
    # images_folder = Path(r'C:\Users\yourname\Documents\job_folder\fotos')

    # # Folder where fotos_renamed/ will be created with the renamed images
    # output_folder = Path(r'C:\Users\yourname\Documents\job_folder\output')
    # ── End of configuration ──────────────────────────────────────────────────

    if not images_folder.exists() or not images_folder.is_dir():
        print(f"ERROR: images folder not found: {images_folder}")
        sys.exit(1)

    output_folder.mkdir(parents=True, exist_ok=True)

    print(f"\nInput  : {images_folder}")
    print(f"Output : {output_folder}")
    print()

    result = extract(images_folder, output_folder)
    print(f"\nDone — {len(result)} images renamed.\n")