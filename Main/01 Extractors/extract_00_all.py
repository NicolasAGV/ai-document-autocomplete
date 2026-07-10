import subprocess
import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / '00'))
import main_path as _mp  # type: ignore

main_path = Path(_mp.main_path)

scripts = [
    "extract_01_sistema.py",
    "extract_02_certificadora.py",
    "extract_03_eut_basic_and_components.py",
    "extract_04_images.py",
    "extract_05_api_ocr_marcado.py",
]

script_dir = Path(__file__).parent
main_dir = str(Path(__file__).parent.parent)
env = os.environ.copy()
env["PYTHONPATH"] = main_dir + (os.pathsep + env["PYTHONPATH"] if "PYTHONPATH" in env else "")

print(f"Running extraction pipeline on: {main_path}\n")

for script in scripts:
    script_path = script_dir / script
    print(f"--- Running {script} ---")
    result = subprocess.run([sys.executable, str(script_path)], cwd=str(script_dir), env=env)
    if result.returncode != 0:
        print(f"ERROR: {script} failed with return code {result.returncode}. Stopping.")
        sys.exit(result.returncode)
    print(f"--- {script} done ---\n")

print("All extraction scripts completed successfully.")
