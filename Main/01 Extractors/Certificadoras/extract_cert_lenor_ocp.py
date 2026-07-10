"""
extract_cert_lenor_ocp.py
=========================
Extracts fields from a Lenor OCP Certificadora PDF and maps them to the
canonical schema defined in extract_02_certificadora.CANONICAL_SCHEMA.

    Canonical key                    ← Lenor OCP source in PDF
    ──────────────────────────────────────────────────────────────────
    Certificadora                    ← hardcoded "Lenor OCP"
    Solicitud de certificadora       ← value to the RIGHT of "Número(s) de proceso"
                                       (words + numbers joined by "-", e.g. "LCSE-455")
    Descripcion del item de ensayo   ← value to the RIGHT of "Producto:"
                                       (e.g. "Fuente de alimentación")
    Identificacion de certificadora  ← value BELOW "Etiqueta N°"  (5-char number, e.g. "42347")
    Marca                            ← value BELOW "Marca"                (e.g. "YC-Friends")
    Modelo                           ← value BELOW "Modelo encontrado"    (e.g. "YC36-2401250")
    Caracteristicas tecnicas         ← value BELOW "Datos o fotos de placa"
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


# Fields that must always be present in a valid Lenor OCP certificadora PDF.
# If any of these comes back empty, extract() emits a warning.
REQUIRED_FIELDS = [
    'Solicitud de certificadora',
    'Identificacion de certificadora',
    'Descripcion del item de ensayo',
]

# "Solicitud de certificadora" looks like word(s) + number(s) joined by "-",
# e.g. "LCSE-455".
_SOLICITUD_RE = re.compile(r'[A-Za-z]+-\d+')


def _clean(value) -> str:
    """Normalise a table cell to a single-line, stripped string."""
    return str(value or '').replace('\n', ' ').strip()


def extract(pdf_path: str | Path) -> dict:
    """Parse a Lenor OCP Certificadora PDF and return the canonical dict."""
    rows = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                rows.extend(table)

    def value_right_of(label: str) -> str:
        """First non-empty cell to the RIGHT of a col-0 label, same row."""
        for row in rows:
            if row and _clean(row[0]) == label:
                for col in range(1, len(row)):
                    val = _clean(row[col])
                    if val:
                        return val
        return ''

    def value_below(label: str) -> str:
        """Value in the data row directly BELOW a header cell.

        Finds the header cell matching `label` on any column, then returns the
        cell in the same column on the next row. Handles spacer/None columns."""
        for i, row in enumerate(rows):
            if not row:
                continue
            for col, cell in enumerate(row):
                if _clean(cell) == label and i + 1 < len(rows):
                    data_row = rows[i + 1]
                    if col < len(data_row):
                        return _clean(data_row[col])
        return ''

    # ── Solicitud de certificadora: right of "Número(s) de proceso" ───────
    solicitud = value_right_of('Número(s) de proceso')
    # Keep only the "LCSE-455"-style token if extra text came along.
    m = _SOLICITUD_RE.search(solicitud)
    if m:
        solicitud = m.group(0)

    # ── Descripcion del item de ensayo: right of "Producto:" ──────────────
    descripcion = value_right_of('Producto:')

    # ── Samples table (values sit BELOW their headers) ────────────────────
    identificacion  = value_below('Etiqueta N°')
    marca           = value_below('Marca')
    modelo          = value_below('Modelo encontrado')
    caracteristicas = value_below('Datos o fotos de placa')

    # Return canonical keys
    result = {
        'Certificadora':                  'Lenor OCP',
        'Solicitud de certificadora':     solicitud,
        'Identificacion de certificadora': identificacion,
        'Descripcion del item de ensayo': descripcion,
        'Marca':                          marca,
        'Modelo':                         modelo,
        'Caracteristicas tecnicas':        caracteristicas,
    }

    # ── Validation: these fields must always be present for Lenor OCP ─────
    missing = [k for k in REQUIRED_FIELDS if not str(result.get(k) or '').strip()]
    if missing:
        warnings.warn(
            "Lenor OCP extraction: required field(s) not found in "
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
