"""
tabla_LC_modelo.py  —  column filler for "tipo/modelo".

TWO ROLES
---------
1. Run directly (python tabla_LC_modelo.py):
   - Reads doc_mod_listado_componentes.json  {name: ["col0", ..., "col4"]}
   - Scans IMAGES_FOLDER for images whose filename contains 'marcado'
     (case-insensitive) and whose stem starts with the component name.
   - Fills slot COL_SLOT with the absolute path to the matching image.
   - Writes the result back to doc_mod_listado_componentes.json.

2. Imported by tabla_LC_nombre.py via FILLER_SCRIPTS:
   - get_data(component_names) reads doc_mod_listado_componentes.json and returns
     {component_name: Path(marcado image)} for every component whose slot
     is non-empty.

Register in tabla_LC_nombre.py:
    FILLER_SCRIPTS = [..., "tabla_LC_modelo", ...]

Matching rule
-------------
  "Capacitor Y 1"  →  prefix "capacitor_y_1"
                       matches "Capacitor_Y_1_Marcado.jpg"  ok
  "Fusible"        →  prefix "fusible"
                       matches "Fusible_Marcado.jpg"         ok
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
COL_SLOT: int = 1

# Folder that contains the component images (including the Marcado ones).
IMAGES_FOLDER: Path = MAIN_FOLDER / "output" / "fotos_renamed"

# OCR results used instead of an image for certain components.
OCR_JSON_PATH: Path = JSON_FOLDER / "extracted_api_ocr_marking.json"

_IMAGE_EXTS   = {".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".tif"}
_trailing_num = re.compile(r"\s+\d+$")


def _strip_accents(s: str) -> str:
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('ascii')

# Component base names that have no marcado image — use OCR text instead.
_OCR_BASES: set[str] = set()

# Component base names that have neither image nor OCR for this column.
_SKIP_BASES = {"capacitor electrolitico"}

# --------------------------------------------------------------------------- #
# OCR lookup                                                                   #
# --------------------------------------------------------------------------- #

# Phrases that mark a line as the OCR model's narration ("thinking") rather
# than an actual marking read off the component. Lines containing / starting
# with these are discarded so only the marking text is kept.
_NARRATION_CUES = (
    "i'm looking", "i am looking", "looking at", "reading all", "reading the",
    "i can see", "i can make out", "i cannot", "i can't", "the image shows",
    "the image appears", "this image", "appears to be", "it appears",
    "would be readable", "with certainty", "the triangle", "the letters",
    "the symbol", "represents", "note:", "unfortunately", "based on",
    "here is", "here's", "the following", "characters on this",
)


def _clean_ocr_text(text: str) -> str:
    """Keep only the marking lines, dropping the model's narration.

    A line is treated as narration (and dropped) when it ends with ':'
    (an intro like "I can see:"), starts with / contains a narration cue
    phrase, or reads as a prose sentence (more than 5 whitespace tokens).
    Short code-like lines ("⚠ PE", "10A 250V~", "YC36-T2-B") are kept."""
    kept: list[str] = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        low = line.lower()
        if line.endswith(":"):
            continue
        if any(cue in low for cue in _NARRATION_CUES):
            continue
        if len(line.split()) > 5:                     # prose sentence
            continue
        kept.append(line)
    return "\n".join(kept).strip()


def _find_ocr_text(component_name: str, ocr_dict: dict[str, str]) -> str | None:
    """Match component name to an OCR dict key by stem prefix.

    Returns only the marking text, with the model's narration removed."""
    plain = _strip_accents(component_name)
    for prefix in (
        plain.replace(" ", "_").lower(),
        _trailing_num.sub("", plain).replace(" ", "_").lower(),
    ):
        for filename, text in ocr_dict.items():
            if Path(filename).stem.lower().startswith(prefix):
                cleaned = _clean_ocr_text(text)
                return cleaned or None
    return None


# --------------------------------------------------------------------------- #
# Image discovery & matching                                                   #
# --------------------------------------------------------------------------- #

def _find_marcado_images(folder: Path) -> list[Path]:
    """Return all images in *folder* whose stem contains 'marcado'."""
    if not folder.is_dir():
        return []
    return [
        p for p in sorted(folder.iterdir())
        if p.suffix.lower() in _IMAGE_EXTS and "marcado" in p.stem.lower()
    ]


def _match_all_marcados(component_name: str, marcado_images: list[Path]) -> list[Path]:
    """
    Return ALL marcado images for *component_name* (sorted).
    Handles both cases:
      - component has number: "Fusible 1"              → prefix "fusible_1_"
      - component has no number (unique): "Interruptor" → matches "Interruptor_1_Marcado…"
        by checking that the image stem starts with the base name followed by _ or end.
    """
    plain = _strip_accents(component_name)
    base  = _trailing_num.sub("", plain)
    norm  = plain.replace(" ", "_").lower()
    bnorm = base.replace(" ", "_").lower()
    return [
        p for p in marcado_images
        if p.stem.lower().startswith(norm) or
           (base == plain and p.stem.lower().startswith(bnorm + "_"))
    ]


# --------------------------------------------------------------------------- #
# Public interface (called by tabla_LC_mod_doc.py)                            #
# --------------------------------------------------------------------------- #

def get_data(component_names: list[str]) -> dict[str, "Path | str"]:
    """
    Return {component_name: value} for components whose COL_SLOT in
    doc_mod_listado_componentes.json is non-empty.
    Value is a Path for image slots, or a str for OCR-text slots.
    """
    if not TEMPLATE_PATH.is_file():
        print(f"  [tabla_LC_modelo] '{TEMPLATE_PATH.name}' not found -- skipping.")
        return {}

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    result: dict[str, "Path | str"] = {}

    for name in component_names:
        base = _trailing_num.sub("", _strip_accents(name)).lower()
        if base in _SKIP_BASES:
            continue
        slots = template.get(name)
        if not slots or COL_SLOT >= len(slots):
            continue
        v = slots[COL_SLOT].strip()
        if not v:
            continue
        p = Path(v)
        if base in _OCR_BASES:
            result[name] = v          # plain text from OCR
        elif p.is_file():
            result[name] = p
        else:
            print(f"  [tabla_LC_modelo] WARNING: image not found for '{name}': {v}")

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

    if not OCR_JSON_PATH.is_file():
        print(f"  Warning: OCR JSON not found: '{OCR_JSON_PATH}' — skipping tipo/modelo.")
        return

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    marcado_images = _find_marcado_images(IMAGES_FOLDER)
    ocr_dict: dict[str, str] = json.loads(OCR_JSON_PATH.read_text(encoding="utf-8"))
    print(f"Marcado images found: {len(marcado_images)}")

    filled = 0
    for name, slots in template.items():
        base = _trailing_num.sub("", _strip_accents(name)).lower()
        if COL_SLOT >= len(slots):
            print(f"  WARNING: '{name}' has only {len(slots)} slot(s) -- skipping.")
            continue

        if base in _SKIP_BASES:
            slots[COL_SLOT] = ""
            print(f"    (skipped: '{name}' has no marcado image)")
            continue

        if base in _OCR_BASES:
            text = _find_ocr_text(name, ocr_dict)
            if text:
                slots[COL_SLOT] = text
                filled += 1
                print(f"    OK  {name} (OCR): {text[:60]}")
            else:
                slots[COL_SLOT] = ""
                print(f"    (no OCR entry for '{name}')")
            continue

        imgs = _match_all_marcados(name, marcado_images)
        if imgs:
            paths_str = ";;".join(str(img) for img in imgs)
            ocr_text  = _find_ocr_text(name, ocr_dict) or ""
            slots[COL_SLOT] = f"{paths_str}||{ocr_text}" if ocr_text else paths_str
            filled += 1
            print(f"    OK  {name}: {len(imgs)} image(s)" + (" + OCR caption" if ocr_text else ""))
        else:
            slots[COL_SLOT] = ""
            print(f"    (no marcado image for '{name}')")

    TEMPLATE_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nUpdated {TEMPLATE_PATH}")
    print(f"  Matched: {filled}/{len(template)} components")


if __name__ == "__main__":
    main()
