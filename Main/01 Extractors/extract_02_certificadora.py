from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

input_folder = Path(_mp.main_path)
output_file = input_folder / 'output' / 'json' / 'extracted_pdf_certificadora.json'

import json
import importlib.util

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run:  pip install pdfplumber")
    sys.exit(1)

_CERTIFICADORAS_DIR = Path(__file__).parent / 'Certificadoras'

# Canonical output schema — all extractors in Certificadoras/ must map to these keys.
CANONICAL_SCHEMA: dict[str, str] = {
    'Certificadora':                  '',
    'Solicitud de certificadora':     '',
    'Identificacion de certificadora': '',
    'Descripcion del item de ensayo': '',
    'Marca':                          '',
    'Modelo':                         '',
    'Caracteristicas tecnicas':        '',
}


def _normalize(data: dict) -> dict:
    """Return a dict guaranteed to contain exactly the canonical keys."""
    return {key: str(data.get(key) or '').strip() for key in CANONICAL_SCHEMA}


def _load_extractor(filename: str):
    """Load a certificadora extractor module from the Certificadoras subfolder."""
    path = _CERTIFICADORAS_DIR / filename
    spec = importlib.util.spec_from_file_location(filename[:-3], path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def detect_certificadora(pdf_path: str | Path) -> str | None:
    """Scan PDF text to identify which certificadora issued it."""
    text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() or ''

    if 'Lenor OCP' in text:
        return 'Lenor OCP'
    if 'Qetkra' in text:
        return 'Qetkra'
    if 'BVE' in text:
        return 'BVE'
    return None


def extract(pdf_path: str | Path) -> dict:
    """Detect certificadora from PDF, delegate to the matching extractor, and
    normalise the result to the canonical schema."""
    pdf_path = Path(pdf_path)
    certificadora = detect_certificadora(pdf_path)

    if certificadora == 'Lenor OCP':
        mod = _load_extractor('extract_cert_lenor_ocp.py')
        raw = mod.extract(pdf_path)

    elif certificadora == 'Qetkra':
        mod = _load_extractor('extract_cert_qetkra.py')
        raw = mod.extract(pdf_path)

    elif certificadora == 'BVE':
        # TODO: implement extract_cert_bve.py
        raise NotImplementedError("Extractor for 'BVE' is not implemented yet.")

    else:
        raise ValueError(f"Could not identify certificadora in: {pdf_path.name}")

    return _normalize(raw)


if __name__ == '__main__':
    candidates = [f for f in input_folder.iterdir()
                  if f.is_file()
                  and f.suffix.lower() == '.pdf'
                  and 'certificadora' in f.name.lower()]

    if not candidates:
        print(f"ERROR: No .pdf file containing 'certificadora' found in: {input_folder}")
        sys.exit(1)

    pdf_file = candidates[0]
    print(f"\nFile: {pdf_file.name}\n")

    try:
        data = extract(pdf_file)
    except NotImplementedError as e:
        print(f"ERROR: {e}")
        sys.exit(1)
    except ValueError as e:
        print(f"ERROR: {e}")
        sys.exit(1)

    print()
    for key, val in data.items():
        print(f"  {key:<35} {val}")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n  ✔ Saved: {output_file}\n")
