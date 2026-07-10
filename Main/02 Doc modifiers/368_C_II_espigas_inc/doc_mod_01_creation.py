import shutil
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent.parent / '00'))
import main_path as _mp  # type: ignore

source_folder      = Path(_mp.main_path)
json_folder        = Path(_mp.main_path) / 'output' / 'json'
output_folder      = Path(_mp.main_path) / 'output'
img_renamed_folder = output_folder / 'fotos_renamed'

normas_folder = Path(__file__).parent.parent.parent.parent / '.XLSX .DOCX patrones' / 'Normas'

EUT_basic_json             = "extracted_xlsx_eut_basic.json"
certificadora_json         = "extracted_pdf_certificadora.json"
componente_existencia_json = "extracted_xlsx_eut_cert_components.json"
sistema_json               = "extracted_htm_sistema.json"


"""
docx_modify_creation.py  —  Part 1
====================================
Selects the .docx template from normas_folder using the 'Norma' value from
extracted_xlsx_eut_basic.json, then copies it to output_folder renamed to
the "Número de Informe" value from extracted_htm_sistema.json.
"""

json_path      = json_folder / sistema_json
eut_basic_path = json_folder / EUT_basic_json

try:
    with open(json_path, encoding='utf-8') as f:
        sistema = json.load(f)

    informe_number = sistema['Número de Informe']

    # ── Resolve the .docx template from the 'Norma' parametro ─────────────────
    with open(eut_basic_path, encoding='utf-8') as f:
        eut_basic = json.load(f)

    norma_value = next(
        (v for k, v in eut_basic.items() if k.strip().lower() == 'norma'),
        None,
    )
    if norma_value is None:
        raise KeyError(f"'Norma' key not found in {eut_basic_path.name}")

    source_docx = normas_folder / f"{norma_value}.docx"
    if not source_docx.is_file():
        raise FileNotFoundError(
            f"Template '{source_docx.name}' not found in Normas folder"
        )

    dest_docx = output_folder / f"{informe_number}.docx"

    output_folder.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source_docx, dest_docx)

    print(f"Norma:  {norma_value}")
    print(f"Copied: {source_docx.name}  →  {dest_docx.name}")

    json_folder.mkdir(parents=True, exist_ok=True)
    informe_json_path = json_folder / "doc_mod_numero_informe.json"
    informe_json_path.write_text(
        json.dumps({"filename": dest_docx.name}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Saved: {informe_json_path}")
except FileNotFoundError as e:
    print(f"Warning: {e} — skipping docx copy.")
except KeyError as e:
    print(f"Warning: {e} — skipping docx copy.")
