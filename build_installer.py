"""Build a standalone executable with PyInstaller including fonts and assets."""
from __future__ import annotations

import os
import PyInstaller.__main__


def build() -> None:
    sep = ';' if os.name == 'nt' else ':'
    datas = [
        f"Header.jpg{sep}.",
        f"Fonts{os.sep}Armata-Regular.ttf{sep}Fonts",
        f"Fonts{os.sep}Novecentowide-Bold.ttf{sep}Fonts",
        f"Fonts{os.sep}Novecentowide-DemiBold_0.ttf{sep}Fonts",
    ]
    PyInstaller.__main__.run([
        "home.py",
        "--name", "ScriptLauncher",
        "--windowed",
        "--onefile",
        *[f"--add-data={d}" for d in datas],
    ])


if __name__ == "__main__":
    build()
