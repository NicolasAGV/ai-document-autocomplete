"""
tabla_LC_nombre.py  —  step 1 of the pipeline.

Run once to filter component names and create the blank template:
    python tabla_LC_nombre.py

What it does
------------
  1. Read extracted_listado_componentes_raw.json → filter and order component names.
  2. Write a fresh doc_mod_listado_componentes.json template:
         {component_name: ["", "", "", "", ""]}   (one slot per data column)
"""

import re
import sys
import json
from pathlib import Path
from collections import defaultdict, Counter

sys.path.insert(0, str(Path(__file__).parent.parent.parent))
import doc_mod_01_creation

# =========================================================================== #
# CONFIGURATION                                                                #
# =========================================================================== #

MASTER_FOLDER = doc_mod_01_creation.source_folder
JSON_FOLDER   = doc_mod_01_creation.json_folder

# Number of data columns after "objeto/parte No." (columns 1-5 in a 6-column table).
N_DATA_COLS: int = 5


# =========================================================================== #
# COMPONENT-NAME LOADING                                                       #
# =========================================================================== #

CATEGORY_ORDER = [
    "Ficha_de_alimentacion",
    "Cordon_de_alimentacion",
    "Conector_de_alimentacion",
    "Zocalo_de_alimentacion",
    "Selector_de_alimentacion",
    "Interruptor",
    "PCB",
    "Fusible",
    "Varistor",
    "Capacitor_X",
    "Bobina",
    "Resistencia",
    "Capacitor_Electrolitico",
    "Capacitor_Y",
    "Optoacoplador",
    "Transformador",
    "Ventilador",
    "Parlante"
]


# Accent corrections — filenames cannot carry tildes, so we restore them here.
_ACCENTS: dict[str, str] = {
    "Electrolitico": "Electrolítico",
    "alimentacion":  "alimentación",
    "Zocalo":        "Zócalo",
    "Cordon":        "Cordón",
}

def _apply_accents(name: str) -> str:
    for plain, accented in _ACCENTS.items():
        name = name.replace(plain, accented)
    return name


_TYPE_SUFFIX = re.compile(r'_(General|Marcado|Logo).*$', re.IGNORECASE)


def load_component_names(json_path: Path) -> list[str]:
    files: list[str] = json.loads(json_path.read_text(encoding="utf-8"))

    # Extract component base name by stripping _General/_Marcado/_Logo suffix.
    # Skip standalone General and Marcado_* images (not components).
    seen: set[str] = set()
    base_stems: list[str] = []
    for f in files:
        stem = Path(f).stem
        base = _TYPE_SUFFIX.sub("", stem)
        if (base.lower().startswith("general")
                or base.lower().startswith("marcado")
                or base.lower().startswith("certificadora")):
            continue
        if base not in seen:
            seen.add(base)
            base_stems.append(base)

    raw_names = [b.replace("_", " ") for b in base_stems]
    _trailing = re.compile(r"\s+\d+$")
    bases      = [_trailing.sub("", n) for n in raw_names]
    counts     = Counter(bases)
    names      = [_apply_accents(b if counts[b] == 1 else r) for r, b in zip(raw_names, bases)]

    # Match against CATEGORY_ORDER using the original unaccented stem so that
    # accented names ("Zócalo…") still find their category ("Zocalo…").
    groups: dict[str, list[str]] = defaultdict(list)
    for name, stem in zip(names, base_stems):
        for cat in CATEGORY_ORDER:
            if stem.startswith(cat):
                groups[cat].append(name)
                break
        else:
            groups["_other"].append(name)

    ordered: list[str] = []
    for cat in CATEGORY_ORDER:
        ordered.extend(groups.get(cat, []))
    ordered.extend(groups.get("_other", []))
    return ordered


# =========================================================================== #
# TEMPLATE EXPORT                                                               #
# =========================================================================== #

def export_component_list(component_names: list[str], out_path: Path) -> None:
    """Write a fresh template {component_name: ["", ..., ""]} to *out_path*."""
    template = {name: [""] * N_DATA_COLS for name in component_names}
    out_path.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Template created → {out_path}")


# =========================================================================== #
# MAIN                                                                         #
# =========================================================================== #

def main() -> None:
    # ------------------------------------------------------------------ #
    # 1. Filter components and create fresh template                       #
    # ------------------------------------------------------------------ #
    raw_path = JSON_FOLDER / "extracted_listado_componentes_raw.json"
    if not raw_path.is_file():
        print(f"Error: '{raw_path}' not found.")
        raise SystemExit(1)

    component_names = load_component_names(raw_path)
    print(f"Components ({len(component_names)}):")
    for n in component_names:
        print(f"  - {n}")
    print()

    JSON_FOLDER.mkdir(parents=True, exist_ok=True)
    template_path = JSON_FOLDER / "doc_mod_listado_componentes.json"
    export_component_list(component_names, template_path)

    print(f"\n{template_path.name}:")
    print(template_path.read_text(encoding="utf-8"))


if __name__ == "__main__":
    main()
