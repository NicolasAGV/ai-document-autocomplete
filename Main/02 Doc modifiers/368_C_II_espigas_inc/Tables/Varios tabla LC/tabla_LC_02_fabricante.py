"""
tabla_LC_fabricante.py  —  column filler for "fabricante/marca registrada".

TWO ROLES
---------
1. Run directly (python tabla_LC_fabricante.py):
   - Reads doc_mod_listado_componentes.json  {name: ["col0", ..., "col4"]}
   - Scans IMAGES_FOLDER for images whose filename contains 'logo'.
   - Matches each logo to a component name.
   - Fills slot COL_SLOT with the absolute path to the logo image for matched
     components, leaves "" for unmatched ones.
   - Writes the result back to doc_mod_listado_componentes.json.

2. Imported by tabla_LC_nombre.py via FILLER_SCRIPTS:
   - get_data(component_names) reads doc_mod_listado_componentes.json and returns
     {component_name: Path(logo image)} for every component whose slot is non-empty.

Register in tabla_LC_nombre.py:
    FILLER_SCRIPTS = ["tabla_LC_fabricante", ...]
"""

import re
import sys
import json
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import doc_mod_01_creation

# --------------------------------------------------------------------------- #
# Configuration                                                                #
# --------------------------------------------------------------------------- #

MAIN_FOLDER   = doc_mod_01_creation.source_folder
JSON_FOLDER   = doc_mod_01_creation.json_folder

TEMPLATE_PATH = JSON_FOLDER / "doc_mod_listado_componentes.json"

# Which slot in the per-component list this filler owns (0-based).
# Columns 1-5: [0]=fabricante, [1]=tipo/modelo, [2]=datos técnicos,
#              [3]=normas,     [4]=marca conformidad
COL_SLOT: int = 0

# Folder to scan for logo images (files whose name contains 'logo').
IMAGES_FOLDER: Path = MAIN_FOLDER / "output" / "fotos_renamed"

_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif", ".webp"}

# --------------------------------------------------------------------------- #
# Logo discovery & matching                                                    #
# --------------------------------------------------------------------------- #

_trailing_num = re.compile(r"\s+\d+$")


def _strip_accents(s: str) -> str:
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('ascii')


def _find_logo_images(folder: Path) -> dict[str, Path]:
    """
    Scan *folder* for images whose stem contains 'logo' (case-insensitive).
    Returns {normalized_component_name: image_path}.

    E.g. "Capacitor_Y_1_Logo.jpg"  → stem "Capacitor_Y_1_Logo"
         strip _Logo               → "Capacitor_Y_1"
         strip trailing _digit     → "Capacitor_Y"
         spaces                    → "Capacitor Y"   ← matches component name
    """
    logo_map: dict[str, Path] = {}
    if not folder.is_dir():
        return logo_map
    for img_path in sorted(folder.iterdir()):
        if img_path.suffix.lower() not in _IMAGE_EXTS:
            continue
        if "logo" not in img_path.stem.lower():
            continue
        norm = re.sub(r"_logo.*$", "", img_path.stem, flags=re.IGNORECASE).strip("_")
        logo_map[norm.replace("_", " ")] = img_path
    return logo_map


def _match_logo(name: str, logo_map: dict[str, Path]) -> Path | None:
    """
    Find a logo for *name*.
    Priority:
      1. Exact match:          "Fusible 1"              → key "Fusible 1"
      2. Base match:           "Fusible 1" base→"Fusible" → key "Fusible"
      3. name starts with key: "Capacitor Y 1"          starts with key "Capacitor Y"
      4. key starts with name: "Ficha de alimentacion"  ← key "Ficha de alimentacion 1"
         (component has no number because it is unique; logo file has the number)
    """
    name_plain = _strip_accents(name)
    base_plain = _strip_accents(_trailing_num.sub("", name))
    name_norm  = name_plain.replace(" ", "_").lower()
    base_norm  = base_plain.replace(" ", "_").lower()

    if name_plain in logo_map:
        return logo_map[name_plain]
    if base_plain in logo_map:
        return logo_map[base_plain]

    # name is longer than key (many-numbered component, single logo)
    for key in sorted(logo_map.keys(), key=len, reverse=True):
        if name_norm.startswith(key.replace(" ", "_").lower()):
            return logo_map[key]

    # key is longer than name (unique component, logo filename has a number)
    for key in sorted(logo_map.keys(), key=len):
        if key.replace(" ", "_").lower().startswith(base_norm):
            return logo_map[key]

    return None


# --------------------------------------------------------------------------- #
# Public interface (called by tabla_LC_mod_doc.py)                            #
# --------------------------------------------------------------------------- #

def get_data(component_names: list[str]) -> dict[str, Path]:
    """
    Return {component_name: Path(logo)} for components whose COL_SLOT in
    doc_mod_listado_componentes.json is non-empty.
    """
    if not TEMPLATE_PATH.is_file():
        print(f"  [tabla_LC_fabricante] '{TEMPLATE_PATH.name}' not found — skipping.")
        return {}

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    result: dict[str, Path] = {}

    for name in component_names:
        slots = template.get(name)
        if not slots or COL_SLOT >= len(slots):
            continue
        v = slots[COL_SLOT].strip()
        if not v:
            continue
        p = Path(v)
        if p.is_file():
            result[name] = p
        else:
            print(f"  [tabla_LC_fabricante] WARNING: logo path not found for '{name}': {v}")

    return result


# --------------------------------------------------------------------------- #
# Standalone: fill COL_SLOT in doc_mod_listado_componentes.json                       #
# --------------------------------------------------------------------------- #

def main() -> None:
    if not TEMPLATE_PATH.is_file():
        print(f"Error: '{TEMPLATE_PATH}' not found.")
        print("Run tabla_LC_nombre.py first to generate doc_mod_listado_componentes.json.")
        raise SystemExit(1)

    if not IMAGES_FOLDER.is_dir():
        print(f"Error: IMAGES_FOLDER not found: '{IMAGES_FOLDER}'.")
        raise SystemExit(1)

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    logo_map = _find_logo_images(IMAGES_FOLDER)

    if not logo_map:
        print(f"  No logo images found in '{IMAGES_FOLDER}'.")

    filled = 0
    for name, slots in template.items():
        if COL_SLOT >= len(slots):
            print(f"  WARNING: '{name}' has only {len(slots)} slot(s) — skipping.")
            continue
        logo_path = _match_logo(name, logo_map)
        if logo_path:
            slots[COL_SLOT] = str(logo_path)
            filled += 1
            print(f"    OK  {name}: {logo_path.name}")
        else:
            slots[COL_SLOT] = ""
            print(f"    (no logo for '{name}')")

    TEMPLATE_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nUpdated {TEMPLATE_PATH}")
    print(f"  Matched: {filled}/{len(template)} components")


if __name__ == "__main__":
    main()
