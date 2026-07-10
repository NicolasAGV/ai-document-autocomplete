from pathlib import Path
import sys
import json
import subprocess

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

input_folder = Path(_mp.main_path)
eut_basic_json = input_folder / 'output' / 'json' / 'extracted_xlsx_eut_basic.json'

_DOC_MODIFIERS_DIR = Path(__file__).parent

NORMA_MAP: dict[str, Path] = {
    '368_C_II_espigas_inc': _DOC_MODIFIERS_DIR / '368_C_II_espigas_inc' / 'doc_mod_00_all.py',
    '950_C_II_espigas_inc': _DOC_MODIFIERS_DIR / '950_C_II_espigas_inc' / 'doc_mod_00_all.py',
}

if not eut_basic_json.exists():
    print(f"ERROR: JSON not found: {eut_basic_json}")
    sys.exit(1)

with open(eut_basic_json, encoding='utf-8') as f:
    eut_data = json.load(f)

norma = str(eut_data.get('Norma') or '').strip()

if not norma:
    print("ERROR: 'Norma' field is empty in extracted_xlsx_eut_basic.json")
    sys.exit(1)

target = NORMA_MAP.get(norma)

if target is None:
    print(f"ERROR: No doc modifier folder mapped for Norma '{norma}'")
    print(f"       Known values: {list(NORMA_MAP)}")
    sys.exit(1)

if not target.exists():
    print(f"ERROR: Target script not found: {target}")
    sys.exit(1)

print(f"Norma: {norma}")
print(f"Running: {target}")

result = subprocess.run([sys.executable, str(target)], check=False)
if result.returncode != 0:
    print(f"\nERROR: {target.name} exited with code {result.returncode}.")
    sys.exit(result.returncode)
