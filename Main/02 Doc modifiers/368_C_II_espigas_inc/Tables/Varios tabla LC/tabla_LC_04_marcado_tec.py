"""
tabla_LC_marcado_tec.py  —  column filler for "Datos técnicos".

TWO ROLES
---------
1. Run directly (python tabla_LC_marcado_tec.py):
   - Reads doc_mod_listado_componentes.json  {name: ["col1", ..., "col5"]}
   - Reads OCR_JSON_PATH  {"Capacitor_Electrolitico_1_Marcado.jpg": "text", ...}
   - Reads PATTERNS_PATH  {"Capacitor Electrolitico": "regex", ...}
   - Extracts technical data for each component (Capacitor Y gets special handling).
   - Fills slot COL_SLOT with the extracted string.
   - Writes the result back to doc_mod_listado_componentes.json.

2. Imported by fill_tabla_custom.py via COLUMN_FILLERS:
   - get_data(component_names) reads doc_mod_listado_componentes.json and returns
     {component_name: str}  for every component whose slot is non-empty.

Register in fill_tabla_custom.py:
    import tabla_LC_marcado_tec
    COLUMN_FILLERS = [..., (tabla_LC_marcado_tec, "datos"), ...]
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
JSON_FOLDER  = doc_mod_01_creation.json_folder

TEMPLATE_PATH = JSON_FOLDER / "doc_mod_listado_componentes.json"

# Columns 1-5 after "objeto/parte No.":
# [0]=fabricante, [1]=tipo/modelo, [2]=datos técnicos, [3]=normas, [4]=marca conformidad
COL_SLOT: int = 2

OCR_JSON_PATH: Path = JSON_FOLDER / "extracted_api_ocr_marking.json"

# _HERE         = Path(__file__).parent          # Tabla de Componentes/
PATTERNS_PATH: Path = Path(__file__).parent / "prompts_datos_tecnicos.json"

# --------------------------------------------------------------------------- #
# OCR helpers                                                                  #
# --------------------------------------------------------------------------- #

def _strip_accents(s: str) -> str:
    return unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode('ascii')


def _find_ocr_text(component_name: str, ocr_dict: dict[str, str]) -> str | None:
    """
    Collect ALL OCR texts whose filename starts with the component prefix
    and return them joined — so multi-image components (e.g. Cordon with
    _Marcado_1 … _Marcado_4) are all searched at once.
    """
    prefix = _strip_accents(component_name).replace(" ", "_").lower()
    texts = [
        text for filename, text in ocr_dict.items()
        if Path(filename).stem.lower().startswith(prefix)
    ]
    return "\n".join(texts) if texts else None

# --------------------------------------------------------------------------- #
# Pattern helpers                                                              #
# --------------------------------------------------------------------------- #

_trailing_num = re.compile(r"\s+\d+$")


def _find_pattern(component_name: str, patterns: dict[str, str]) -> str | None:
    """Exact → base name → longest prefix match. Accent-insensitive."""
    plain = _strip_accents(component_name)
    base  = _trailing_num.sub("", plain)
    if plain in patterns:
        return patterns[plain]
    if base in patterns:
        return patterns[base]
    for key in sorted(patterns.keys(), key=len, reverse=True):
        if plain.replace(" ", "_").lower().startswith(key.replace(" ", "_").lower()):
            return patterns[key]
    return None


def _normalize_units(text: str) -> str:
    """Replace OCR artefact 'uF' / 'UF' with the correct symbol 'μF'."""
    return re.sub(r"[Uu][Ff]", "μF", text)


def _extract(ocr_text: str, pattern: str) -> str | None:
    m = re.search(pattern, ocr_text, re.IGNORECASE)
    return _normalize_units(m.group(0).strip()) if m else None

# --------------------------------------------------------------------------- #
# Capacitor Y — code-based capacitance extraction                             #
# --------------------------------------------------------------------------- #

_CAP_Y_VALID_CODES = ["221", "471", "222", "332", "472", "101", "102"]


def _hamming(a: str, b: str) -> int:
    return sum(x != y for x, y in zip(a, b))


def _closest_cap_y_code(raw: str) -> str:
    return min(_CAP_Y_VALID_CODES, key=lambda c: _hamming(raw, c))


def _decode_cap_y(code: str) -> str:
    pf = int(code[:2]) * (10 ** int(code[2]))
    nf = pf / 1000
    if nf == int(nf):
        return f"{int(nf)}nF"
    return f"{nf:.2f}".rstrip("0").replace(".", ",") + "nF"


def _extract_capacitor_y(ocr_text: str) -> str | None:
    """
    Extract the 3 required markings from a Capacitor Y:
      1. Capacitance — 3-digit code (fuzzy-matched) decoded to nF, e.g. 3,3 nF
      2. Voltage     — one of the standard Y-cap ratings: 125/250/275/300/400/440/500 V
      3. Insulation class — Y1 or Y2

    Returns a string combining whatever was found (e.g. "3,3 nF Y1 400V").
    Returns None if none of the three could be extracted.
    """
    # 1. Capacitance — exclude numbers that look like voltages (125, 250, 300, 400 …)
    _voltage_vals = {"125", "250", "275", "300", "400", "440", "500"}
    candidates = [c for c in re.findall(r"\d{3}", ocr_text) if c not in _voltage_vals]
    if candidates:
        best_raw = min(candidates, key=lambda r: min(_hamming(r, c) for c in _CAP_Y_VALID_CODES))
        code = best_raw if best_raw in _CAP_Y_VALID_CODES else _closest_cap_y_code(best_raw)
        capacity = _decode_cap_y(code)
    else:
        capacity = None

    # 2. Voltage — standard Y-capacitor AC ratings; unit is V or ~ (AC tilde)
    v_match = re.search(r"(125|250|275|300|400|440|500)\s*(?:[Vv]|~)", ocr_text)
    if v_match:
        digits  = v_match.group(1)
        voltage = digits + "V"   # normalise ~ → V in output
    else:
        voltage = None

    # 3. Insulation class — Y1/Y2 can appear anywhere in the string,
    #    including embedded in longer tokens like "Y1250~" or "TYPE-Y2X"
    y_match = re.search(r"Y[12]", ocr_text, re.IGNORECASE)
    y_class = y_match.group(0).upper() if y_match else None

    parts = [p for p in [capacity, y_class, voltage] if p is not None]
    return " ".join(parts) if parts else None

# --------------------------------------------------------------------------- #
# Fusible — nomenclature-based extraction                                     #
# --------------------------------------------------------------------------- #

def _extract_fusible(ocr_text: str) -> str | None:
    """
    Extract fuse markings of the form T1.6AL250V (IEC 60127 style):
      - Leading letter: T or F (time-lag / fast-blow)
      - Current: decimal number + A  (e.g. 1.6A, 2A)
      - Type letter: L or H
      - Voltage: 2-4 digit number + V  (e.g. 250V, 125V)
    Falls back to returning current + voltage separately if the full
    pattern is not found.
    """
    # Full IEC-style marking
    m = re.search(r"[TF]\d+(?:[.,]\d+)?A[LH]\d{2,4}V", ocr_text, re.IGNORECASE)
    if m:
        return m.group(0).upper()

    # Fallback: extract current and voltage individually
    parts = []
    a_match = re.search(r"\d+(?:[.,]\d+)?\s*A\b", ocr_text, re.IGNORECASE)
    v_match = re.search(r"\d{2,4}\s*V\b", ocr_text, re.IGNORECASE)
    if a_match:
        parts.append(a_match.group(0).replace(" ", "").upper())
    if v_match:
        parts.append(v_match.group(0).replace(" ", "").upper())
    return " ".join(parts) if parts else None

# --------------------------------------------------------------------------- #
# Voltage + Current — generic V/A extraction                                  #
# --------------------------------------------------------------------------- #

# Matches voltage in all common forms:
#   250V  250V~  ~250V  200/500V  200/500V~  250~
# Handles slash-separated ranges and tilde as AC symbol before or after.
_V_RE = re.compile(
    r"~?\d+(?:[.,]\d+)?(?:/\d+(?:[.,]\d+)?)?\s*(?:V[~]?|~(?!\d))",
    re.IGNORECASE,
)

# Matches cable section: 0.75mm²  1.5mm2  0,75 mm²  .75mm² (OCR drops leading 0)
_MM2_RE = re.compile(r"(?:\d+(?:[.,]\d+)?|[.,]\d+)\s*mm\s*[²2]", re.IGNORECASE)

# Components whose marking contains voltage (V) and current (A).
_VA_BASES = {
    "ficha de alimentacion",
    "conector de alimentacion",
    "zocalo de alimentacion",
    "selector de alimentacion",
    "interruptor",
    "ventilador",
}


def _fmt_voltage(m: re.Match) -> str:
    """Normalise a voltage match: strip spaces, uppercase."""
    return m.group(0).replace(" ", "").upper()


def _fmt_mm2(m: re.Match) -> str:
    """Normalise a mm² match: strip spaces, replace trailing 2 with ²."""
    raw = m.group(0).replace(" ", "")
    return re.sub(r"mm2$", "mm²", raw, flags=re.IGNORECASE)


def _extract_voltage_current(ocr_text: str) -> str | None:
    """Extract voltage (V) and current (A), return e.g. '200/500V~ 10A'."""
    v_match = _V_RE.search(ocr_text)
    a_match = re.search(r"\d+(?:[.,]\d+)?\s*A\b", ocr_text, re.IGNORECASE)
    parts = []
    if v_match:
        parts.append(_fmt_voltage(v_match))
    if a_match:
        parts.append(a_match.group(0).replace(" ", "").upper())
    return " ".join(parts) if parts else None


# --------------------------------------------------------------------------- #
# Parlante — impedance extraction                                              #
# --------------------------------------------------------------------------- #

def _extract_parlante(ocr_text: str) -> str | None:
    """Extract impedance in Ω, e.g. '8Ω'."""
    m = re.search(r"\d+(?:[.,]\d+)?\s*[ΩΩ]", ocr_text)
    if m:
        return m.group(0).replace(" ", "")
    m = re.search(r"\d+(?:[.,]\d+)?\s*ohm", ocr_text, re.IGNORECASE)
    return m.group(0).replace(" ", "").lower() if m else None


# --------------------------------------------------------------------------- #
# Cordon de alimentacion — voltage + cable section                             #
# --------------------------------------------------------------------------- #

def _extract_cordon(ocr_text: str) -> str | None:
    """Extract voltage (V) and cable section (mm²), e.g. '200/500V~ 0.75mm²'."""
    v_match   = _V_RE.search(ocr_text)
    mm2_match = _MM2_RE.search(ocr_text)
    parts = []
    if v_match:
        parts.append(_fmt_voltage(v_match))
    if mm2_match:
        parts.append(_fmt_mm2(mm2_match))
    return " ".join(parts) if parts else None


# --------------------------------------------------------------------------- #
# Core extraction                                                              #
# --------------------------------------------------------------------------- #

def _extract_all(component_names: list[str],
                 ocr_dict: dict[str, str],
                 patterns: dict[str, str]) -> dict[str, str]:
    extracted: dict[str, str] = {}
    for name in component_names:
        ocr_text = _find_ocr_text(name, ocr_dict)
        if ocr_text is None:
            print(f"    (no OCR entry for '{name}')")
            continue

        base = _strip_accents(_trailing_num.sub("", name)).lower()

        if base == "transformador":
            print(f"    (skipped: Transformador has no technical marking)")
            continue

        if base in _VA_BASES:
            result = _extract_voltage_current(ocr_text)
            if result:
                extracted[name] = result
                print(f"    OK  {name}: {result}")
            else:
                print(f"    (no V/A values found in OCR for '{name}')")
            continue

        if base == "parlante":
            result = _extract_parlante(ocr_text)
            if result:
                extracted[name] = result
                print(f"    OK  {name}: {result}")
            else:
                print(f"    (no impedance found in OCR for '{name}')")
            continue

        if base == "cordon de alimentacion":
            result = _extract_cordon(ocr_text)
            if result:
                extracted[name] = result
                print(f"    OK  {name}: {result}")
            else:
                print(f"    (no V/mm² values found in OCR for '{name}')")
            continue

        if base == "capacitor y":
            result = _extract_capacitor_y(ocr_text)
            if result:
                extracted[name] = result
                print(f"    OK  {name}: {result}")
            else:
                print(f"    (no 3-digit code in OCR for '{name}')")
            continue

        if base == "fusible":
            result = _extract_fusible(ocr_text)
            if result:
                extracted[name] = result
                print(f"    OK  {name}: {result}")
            else:
                print(f"    (no fuse marking found in OCR for '{name}')")
            continue

        pattern = _find_pattern(name, patterns)
        if pattern is None:
            print(f"    (no pattern for '{name}')")
            continue
        result = _extract(ocr_text, pattern)
        if result:
            extracted[name] = result
            print(f"    OK  {name}: {result}")
        else:
            print(f"    (pattern no match for '{name}')")
    return extracted

# --------------------------------------------------------------------------- #
# Public interface (called by fill_tabla_custom.py)                           #
# --------------------------------------------------------------------------- #

def get_data(component_names: list[str]) -> dict[str, str]:
    """
    Return {component_name: str} for components whose COL_SLOT in
    doc_mod_listado_componentes.json is non-empty.
    """
    if not TEMPLATE_PATH.is_file():
        print(f"  [tabla_LC_marcado_tec] '{TEMPLATE_PATH.name}' not found — skipping.")
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
        print("Run fill_tabla_custom.py first to generate doc_mod_listado_componentes.json.")
        raise SystemExit(1)
    if not OCR_JSON_PATH.is_file():
        print(f"  Warning: OCR JSON not found: '{OCR_JSON_PATH}' — skipping datos técnicos.")
        return
    if not PATTERNS_PATH.is_file():
        print(f"Error: patterns JSON not found: '{PATTERNS_PATH}'.")
        raise SystemExit(1)

    template: dict[str, list[str]] = json.loads(TEMPLATE_PATH.read_text(encoding="utf-8"))
    ocr_dict: dict[str, str]       = json.loads(OCR_JSON_PATH.read_text(encoding="utf-8"))
    patterns: dict[str, str]       = json.loads(PATTERNS_PATH.read_text(encoding="utf-8"))

    component_names = list(template.keys())
    print(f"Components: {len(component_names)}")

    extracted = _extract_all(component_names, ocr_dict, patterns)

    for name, slots in template.items():
        if COL_SLOT >= len(slots):
            print(f"  WARNING: '{name}' has only {len(slots)} slot(s) — skipping.")
            continue
        slots[COL_SLOT] = extracted.get(name, "")

    TEMPLATE_PATH.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\nUpdated {TEMPLATE_PATH}")
    print(f"  Filled: {len(extracted)}/{len(component_names)} components")


if __name__ == "__main__":
    main()
