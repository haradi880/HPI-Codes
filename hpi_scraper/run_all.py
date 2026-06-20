from __future__ import annotations

import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parent

SCRIPTS = [
    "build_firmographic_structure.py",
    "build_jobs_hiring_structure.py",
    "build_technographic_structure.py",
    "news_announcements_fetch.py",
    "build_contact_level_structure.py",
]


def main() -> None:
    for script in SCRIPTS:
        print(f"\n=== Running {script} ===", flush=True)
        subprocess.run([sys.executable, str(ROOT / script)], cwd=ROOT, check=True)


if __name__ == "__main__":
    main()
