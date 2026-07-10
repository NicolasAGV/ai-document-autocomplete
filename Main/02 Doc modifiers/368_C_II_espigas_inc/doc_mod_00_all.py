import runpy
from pathlib import Path

"""
docx_modify_all.py
===================
Runs the full .docx build pipeline in order:
    1. docx_modify_creation  — copy and rename the template
    2. docx_modify_markings  — fill text fields
    3. docx_modify_clauses   — apply clause logic
    4. docx_modify_T_LC      — fill Listado de Componentes table
    5. docx_modify_images    — insert images
"""

here = Path(__file__).parent

scripts = [
    'doc_mod_01_creation.py',
    'doc_mod_02_markings.py',
    'doc_mod_03_clauses.py',
    'doc_mod_04_tables.py',
    'doc_mod_05_images.py',
]

for script in scripts:
    print(f"\n{'='*60}")
    print(f"  {script}")
    print(f"{'='*60}")
    runpy.run_path(str(here / script), run_name='__main__')

print("\nAll done.\n")
