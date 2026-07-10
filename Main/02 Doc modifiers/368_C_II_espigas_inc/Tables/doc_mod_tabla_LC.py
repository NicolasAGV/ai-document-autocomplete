# EN MODO STAND ALONE: agregar "sistema.htm" en la carpeta raiz

import runpy
from pathlib import Path

"""
docx_modify_T_LC.py  —  Part 3
================================
Runs the full Listado de Componentes pipeline in order:
    1. tabla_LC_nombre      — filter component names, build blank template
    2. tabla_LC_norma       — fill Norma column
    3. tabla_LC_modelo      — fill Modelo column
    4. tabla_LC_marcado_tec — fill Marcado técnico column
    5. tabla_LC_fabricante  — fill Fabricante column
    6. tabla_LC_certificado — fill Certificado column
    7. tabla_LC_mod_doc     — write filled table into the .docx
"""

tabla_folder = Path(__file__).parent

scripts = [
    'Varios tabla LC/tabla_LC_01_nombre.py',
    'Varios tabla LC/tabla_LC_02_fabricante.py',
    'Varios tabla LC/tabla_LC_03_modelo.py',
    'Varios tabla LC/tabla_LC_04_marcado_tec.py',
    'Varios tabla LC/tabla_LC_05_norma.py',        
    'Varios tabla LC/tabla_LC_06_certificado.py',
    'Varios tabla LC/tabla_LC_07_mod_doc.py',
]

for script in scripts:
    print(f"\n  — {script} —")
    runpy.run_path(str(tabla_folder / script), run_name='__main__')

print("\nDone.\n")
