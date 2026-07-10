import runpy
from pathlib import Path

here = Path(__file__).parent

scripts = [
    here / 'extract_images.py',
    here / 'extract_APIOCR_marcado.py',
    here / 'Tables' / 'Tabla_de_componentes' / 'docx_modify_T_LC.py',
    here / 'docx_modify_images.py',
]

for script in scripts:
    print(f"\n{'='*60}\n  — {script.name} —\n{'='*60}")
    runpy.run_path(str(script), run_name='__main__')

print("\nDone.\n")
