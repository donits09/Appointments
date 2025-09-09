
"""Build standalone executables for the suite using PyInstaller."""

from __future__ import annotations

import os
import PyInstaller.__main__


SEP = ';' if os.name == 'nt' else ':'
DATAS = [
    f"Header.jpg{SEP}.",
    f"Fonts{os.sep}Armata-Regular.ttf{SEP}Fonts",
    f"Fonts{os.sep}Novecentowide-Bold.ttf{SEP}Fonts",
    f"Fonts{os.sep}Novecentowide-DemiBold_0.ttf{SEP}Fonts",
    f"favicon.ico{SEP}.",
]


def build_one(script: str, name: str) -> None:
    """Invoke PyInstaller for ``script`` producing ``name``."""
    PyInstaller.__main__.run([
        script,
        "--name",
        name,
        "--windowed",
        "--onefile",
        "--distpath",
        "dist",
        "--workpath",
        f"build/{name}",
        *[f"--add-data={d}" for d in DATAS],
    ])


def build() -> None:
    build_one("home.py", "ScriptLauncher")
    build_one("Appointments/main_v2.py", "Appointments")
    build_one("Payments/main_v2.py", "Payments")
    build_one("Pending/main_v2.py", "Pending")
    build_one("pdf.py", "PDFViewer")


if __name__ == "__main__":
    build()
