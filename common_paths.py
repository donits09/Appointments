from __future__ import annotations

"""Shared helper functions for locating user data files."""

from pathlib import Path
import os
import sys

APP_NAME = "ScriptLauncher"


def base_dir() -> Path:
    """Return directory containing this application or executable."""
    if getattr(sys, "frozen", False):
        # When bundled with PyInstaller ``sys._MEIPASS`` points to the
        # temporary extraction directory that holds bundled data files.  Use
        # it when available so relative resources like fonts are found both in
        # one-file and one-folder distributions.
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    """Return the directory where user data (CSV files) is stored.

    Creates the directory if it does not already exist.
    """
    root_dir = Path(os.getenv("LOCALAPPDATA", base_dir()))
    d = root_dir / APP_NAME / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def find_latest_csv(prefix: str) -> Path | None:
    """Return the most recently modified CSV file starting with ``prefix``.

    Parameters
    ----------
    prefix:
        Filename prefix to match, e.g. ``"appointments_"``.

    Returns
    -------
    Path | None
        Path to the latest matching CSV file or ``None`` if none exist.
    """
    candidates = sorted(
        data_dir().glob(f"{prefix}*.csv"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None
