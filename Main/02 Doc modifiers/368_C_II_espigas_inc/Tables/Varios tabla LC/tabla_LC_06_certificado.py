"""
tabla_LC_marca.py  —  column filler for "marca(s) de conformidad1)".

TWO ROLES
---------
1. Run directly (python tabla_LC_marca.py):
   - Reads doc_mod_listado_componentes.json  {name: ["col1", "col2", ..., "col5"]}
   - Checks each component against extracted_xlsx_eut_cert_components.json
   - Fills slot COL_SLOT with the absolute path to Logo UL.jpg for certified
     components, leaves "" for uncertified ones.
   - Writes the result back to doc_mod_listado_componentes.json.

2. Imported by fill_tabla_custom.py via COLUMN_FILLERS:
   - get_data(component_names) reads doc_mod_listado_componentes.json and returns
     {component_name: Path(Logo UL.jpg)} for every component whose slot
     COL_SLOT is non-empty.

Register in fill_tabla_custom.py:
    import tabla_LC_marca
    COLUMN_FILLERS = [..., (tabla_LC_marca, "conformi")]
"""

import re
import sys
import json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import doc_mod_01_creation

# --------------------------------------------------------------------------- #
# Configuration                                                                #
# --------------------------------------------------------------------------- #

MAIN_FOLDER   = doc_mod_01_creation.source_folder
JSON_FOLDER   = doc_mod_01_creation.json_folder

TEMPLATE_PATH = JSON_FOLDER / "doc_mod_listado_componentes.json"

# Which slot in the per-component list this filler owns (0-based, after col 0).
# Columns 1-5: [0]=fabricante, [1]=tipo/modelo, [2]=datos técnicos,
#              [3]=normas,     [4]=marca(s) de conformidad
COL_SLOT: int = 4

# JSON listing certified component base-names.
# Format: ["Capacitor Electrolitico", "PCB", ...]  or {"Name": ..., ...}
CERTIFIED_PATH: Path = MAIN_FOLDER / "output" / "json" / "extracted_xlsx_eut_cert_components.json"

# UL logo image — must sit in the same folder as this script.
UL_LOGO_PATH: Path = Path(__file__).parent / "Logo UL.jpg"

# --------------------------------------------------------------------------- #
# Internal helpers                                                             #
# --------------------------------------------------------------------------- #

_trailing_num = re.compile(r"\s+\d+$")


def _load_certified(path: Path) -> set[str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    names = data if isinstance(data, list) else list(data.keys())
    return {str(n).strip().lower() for n in names}


def _is_certified(name: str, certified: set[str]) -> bool:
    base = _trailing_num.sub("", name).strip().lower()
    return base in certified or name.strip().lower() in certified


# --------------------------------------------------------------------------- #
# Public interface (called by fill_tabla_custom.py)                           #
# --------------------------------------------------------------------------- #

def get_data(component_names: list[str]) -> dict[str, Path]:
    """
    Return {component_name: Path(UL logo)} for components whose COL_SLOT
    in doc_mod_listado_componentes.json is non-empty.
    """
    if not TEMPLATE_PATH.is_file():
        print(f"  [tabla_LC_marca] '{TEMPLATE_PATH.name}' not found — skipping.")
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
            print(f"  [tabla_LC_marca] WARNING: logo path not found for '{name}': {v}")

    return result


# --------------------------------------------------------------------------- #
# Standalone: fill COL_SLOT in doc_mod_listado_componentes.json                       #
# --------------------------------------------------------------------------- #

def main() -> None:
    if not TEMPLATE_PATH.is_file():
        print(f"Error: '{TEMPLATE_PATH}' not found.")
        print("Run fill_tabla_custom.py first to generate doc_mod_listado_componentes.json.")
        raise SystemExit(1)

    if not CERTIFIED_PATH.is_file():
        print(f"  Warning: certified-components file not found: '{CERTIFIED_PATH}' — skipping marca conformidad.")
        return

    if not UL_LOGO_PATH.is_file():
        print(f"Error: UL logo not found: '{UL_LOGO_PATH}'.")
        raise SystemExit(1)

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    certified = _load_certified(CERTIFIED_PATH)

    filled = 0
    for name, slots in template.items():
        if COL_SLOT >= len(slots):
            print(f"  WARNING: '{name}' has only {len(slots)} slot(s), expected > {COL_SLOT} — skipping.")
            continue
        if _is_certified(name, certified):
            slots[COL_SLOT] = str(UL_LOGO_PATH)
            filled += 1
        else:
            slots[COL_SLOT] = ""

    TEMPLATE_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Updated {TEMPLATE_PATH}")
    print(f"  Certified (logo set): {filled}/{len(template)} components")
    for name, slots in template.items():
        tag = f"OK  {slots[COL_SLOT]}" if slots[COL_SLOT] else "(not certified)"
        print(f"  {name}: {tag}")


if __name__ == "__main__":
    main()
