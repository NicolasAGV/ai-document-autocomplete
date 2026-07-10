import subprocess
import sys
from pathlib import Path

base = Path(__file__).parent.parent

scripts = [
    base / "01 Extractors/extract_00_all.py",
    base / "02 Doc modifiers/standard_selection.py",
]

for script in scripts:
    print(f"\n{'='*60}")
    print(f"Running: {script.name}")
    print('='*60)
    result = subprocess.run([sys.executable, str(script)], check=False)
    if result.returncode != 0:
        print(f"\nERROR: {script.name} exited with code {result.returncode}. Stopping.")
        sys.exit(result.returncode)

print("\nAll stages completed successfully.")
