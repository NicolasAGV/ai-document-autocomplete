"""
tabla_LC_norma.py  —  column filler for "normas".

TWO ROLES
---------
1. Run directly (python tabla_LC_norma.py):
   - Reads listado_componentes.json  {name: ["col0", ..., "col4"]}
   - Looks up each component name in dict_norma_script.json.
   - Fills slot COL_SLOT with the matching norm string.
   - Writes the result back to listado_componentes.json.

2. Imported by tabla_LC_nombre.py via FILLER_SCRIPTS:
   - get_data(component_names) reads doc_mod_listado_componentes.json and returns
     {component_name: str} for every component whose slot is non-empty.

Register in tabla_LC_nombre.py:
    FILLER_SCRIPTS = [..., "tabla_LC_norma", ...]

dict_norma_script.json format:
    { "Capacitor X": "UL 60384-14", "Fusible": "UL 248-14", ... }
Keys are matched exactly, then by base name (trailing number stripped).
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
COL_SLOT: int = 3

_HERE          = Path(__file__).parent          # Tabla de Componentes/
NORMA_DICT_PATH: Path = _HERE / "dict_norma_script.json"

# --------------------------------------------------------------------------- #
# Lookup helpers                                                               #
# --------------------------------------------------------------------------- #

_trailing_num = re.compile(r"\s+\d+$")


def _nfc(s: str) -> str:
    return unicodedata.normalize("NFD", s).encode("ascii", "ignore").decode("ascii").lower()


def _lookup(name: str, norma_dict: dict[str, str]) -> str | None:
    """Exact match → base name → accent-insensitive match."""
    if name in norma_dict:
        return norma_dict[name]
    base = _trailing_num.sub("", name)
    if base in norma_dict:
        return norma_dict[base]
    # Fallback: compare after NFC normalisation + lowercase so that
    # accented names ("Cordón de alimentación") match unaccented dict
    # keys ("Cordon de alimentacion") and vice-versa.
    norm_map = {_nfc(k): v for k, v in norma_dict.items()}
    return norm_map.get(_nfc(name)) or norm_map.get(_nfc(base))


# --------------------------------------------------------------------------- #
# Public interface (called by tabla_LC_mod_doc.py)                            #
# --------------------------------------------------------------------------- #

def get_data(component_names: list[str]) -> dict[str, str]:
    """
    Return {component_name: str} for components whose COL_SLOT in
    doc_mod_listado_componentes.json is non-empty.
    """
    if not TEMPLATE_PATH.is_file():
        print(f"  [tabla_LC_norma] '{TEMPLATE_PATH.name}' not found — skipping.")
        return {}

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    result: dict[str, str] = {}

    for name in component_names:
        slots = template.get(name)
        if slots and COL_SLOT < len(slots) and slots[COL_SLOT].strip():
            result[name] = slots[COL_SLOT].strip()

    return result


# --------------------------------------------------------------------------- #
# Standalone: fill COL_SLOT in doc_mod_listado_componentes.json                       #
# --------------------------------------------------------------------------- #

def main() -> None:
    if not TEMPLATE_PATH.is_file():
        print(f"Error: '{TEMPLATE_PATH}' not found.")
        print("Run tabla_LC_nombre.py first to generate doc_mod_listado_componentes.json.")
        raise SystemExit(1)

    if not NORMA_DICT_PATH.is_file():
        print(f"Error: norma dict not found: '{NORMA_DICT_PATH}'.")
        raise SystemExit(1)

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    norma_dict: dict[str, str]     = json.loads(NORMA_DICT_PATH.read_text(encoding="utf-8"))

    filled = 0
    for name, slots in template.items():
        if COL_SLOT >= len(slots):
            print(f"  WARNING: '{name}' has only {len(slots)} slot(s) — skipping.")
            continue
        norma = _lookup(name, norma_dict)
        if norma:
            slots[COL_SLOT] = norma
            filled += 1
            print(f"    OK  {name}: {norma}")
        else:
            slots[COL_SLOT] = ""
            print(f"    (no norma for '{name}')")

    TEMPLATE_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nUpdated {TEMPLATE_PATH}")
    print(f"  Filled: {filled}/{len(template)} components")


if __name__ == "__main__":
    main()
