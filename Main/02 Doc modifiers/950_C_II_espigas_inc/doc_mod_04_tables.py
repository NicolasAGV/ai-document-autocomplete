import runpy
from pathlib import Path

"""
docx_modify_tables.py
=====================
Runs all table-modification scripts in order:
    1. Tables\Tabla_de_componentes\docx_modify_T_LC.py  — Listado de Componentes
    2. Tables\tabla_condiciones_ensayo.py               — Condiciones de Ensayo
"""

_HERE = Path(__file__).parent

scripts = [
    _HERE / 'Tables' / 'doc_mod_tabla_LC.py',
    _HERE / 'Tables' / 'doc_mod_tabla_cond_ensayo.py',
    _HERE / 'Tables' / 'doc_mod_tabla_espigas.py',
]

for script in scripts:
    print(f"\n  ══ {script.name} ══")
    runpy.run_path(str(script), run_name='__main__')

print("\nDone.\n")
