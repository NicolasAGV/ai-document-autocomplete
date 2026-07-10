"""
extract_cert_qetkra.py
=======================
Extracts fields from a Qetkra Certificadora PDF and maps them to the
canonical schema defined in extract_02_certificadora.CANONICAL_SCHEMA.

    Canonical key                    ← Qetkra source in PDF
    ──────────────────────────────────────────────────────────────────
    Certificadora                    ← hardcoded "Qetkra"
    Solicitud de certificadora       ← value to the RIGHT of "Nro. de Certificado / Proceso"
                                       (e.g. "Q26-00942-01")
    Identificacion de certificadora  ← value to the RIGHT of "Nro. de Lacre"
                                       (e.g. "D-891")
    Descripcion del item de ensayo   ← value to the RIGHT of "Descripción"
                                       (e.g. "Fuente de alimentación para PC")
    Marca                            ← value to the RIGHT of "Marca"          (e.g. "NOXI")
    Modelo                           ← value to the RIGHT of "Modelo a Ensayar"
                                       (e.g. "NPSG-1000 / NOX-49")
    Caracteristicas tecnicas         ← value to the RIGHT of "Características Técnicas"

Unlike the Lenor OCP PDFs, this document has no extractable tables — the
"label | value" layout you see visually is just plain text with the label
followed by the value on the same line (pdfplumber's extract_tables()
returns nothing for this PDF). So we parse the raw text line-by-line
instead, matching each label at the start of a line and taking everything
after it as the value.
"""

from pathlib import Path
import re
import sys
import warnings

sys.path.insert(0, str(Path(__file__).parent.parent))
import extract_02_certificadora as _cert02  # type: ignore

input_folder = _cert02.input_folder

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run:  pip install pdfplumber")
    sys.exit(1)


# Fields that must always be present in a valid Qetkra certificadora PDF.
# If any of these comes back empty, extract() emits a warning.
REQUIRED_FIELDS = [
    'Solicitud de certificadora',
    'Identificacion de certificadora',
    'Descripcion del item de ensayo',
]


def _clean(value) -> str:
    """Normalise extracted text to a single-line, stripped string."""
    return re.sub(r'\s+', ' ', str(value or '')).strip()


def _value_right_of(text: str, label: str) -> str:
    """Return the text following `label` on the same line, up to line end.

    Matches `label` at the start of a line (labels in this PDF never repeat
    mid-sentence) and returns everything after it, stripped."""
    pattern = re.compile(
        r'^' + re.escape(label) + r'\s*(.*)$',
        re.MULTILINE,
    )
    m = pattern.search(text)
    return _clean(m.group(1)) if m else ''


def extract(pdf_path: str | Path) -> dict:
    """Parse a Qetkra Certificadora PDF and return the canonical dict."""
    text_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ''
            text_parts.append(page_text)
    text = '\n'.join(text_parts)

    solicitud       = _value_right_of(text, 'Nro. de Certificado / Proceso')
    identificacion  = _value_right_of(text, 'Nro. de Lacre')
    descripcion     = _value_right_of(text, 'Descripción')
    marca           = _value_right_of(text, 'Marca')
    modelo          = _value_right_of(text, 'Modelo a Ensayar')
    caracteristicas = _value_right_of(text, 'Características Técnicas')

    # Return canonical keys
    result = {
        'Certificadora':                  'Qetkra',
        'Solicitud de certificadora':     solicitud,
        'Identificacion de certificadora': identificacion,
        'Descripcion del item de ensayo': descripcion,
        'Marca':                          marca,
        'Modelo':                         modelo,
        'Caracteristicas tecnicas':        caracteristicas,
    }

    # ── Validation: these fields must always be present for Qetkra ────────
    missing = [k for k in REQUIRED_FIELDS if not str(result.get(k) or '').strip()]
    if missing:
        warnings.warn(
            "Qetkra extraction: required field(s) not found in "
            f"'{Path(pdf_path).name}': {', '.join(missing)}",
            stacklevel=2,
        )

    return result


if __name__ == '__main__':
    candidates = [f for f in input_folder.iterdir()
                  if f.is_file()
                  and f.suffix.lower() == '.pdf'
                  and 'certificadora' in f.name.lower()]

    if not candidates:
        print(f"ERROR: No .pdf containing 'certificadora' found in: {input_folder}")
        sys.exit(1)

    pdf_file = candidates[0]
    print(f"\nFile: {pdf_file.name}\n")

    data = extract(pdf_file)
    for key, val in data.items():
        print(f"  {key:<40} {val}")