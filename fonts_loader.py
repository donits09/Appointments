from __future__ import annotations

"""Utilities for registering bundled font files with the OS/Tk.

The original application imports :func:`load_private_fonts` before any
``tkinter`` widgets are created.  The function therefore avoids using
``tkinter`` directly and instead relies on platform specific mechanisms for
making fonts available.  Only the behaviour needed by the project is
implemented here which keeps the function intentionally small and without
side effects when run on unsupported platforms.
"""

from pathlib import Path
import sys
from typing import Iterable

try:  # pragma: no cover - optional dependency on Windows
    import ctypes
except Exception:  # pragma: no cover - handled gracefully
    ctypes = None


def _add_windows_font(font_path: Path) -> None:
    """Register *font_path* with Windows using ``AddFontResourceEx``.

    The call is best-effort; any failure is silently ignored so that the
    application continues to run even when fonts cannot be registered.  This
    mirrors the permissive behaviour of the original project.
    """

    if ctypes is None:
        return

    FR_PRIVATE = 0x10  # flag to keep the font private to the process
    try:
        ctypes.windll.gdi32.AddFontResourceExW(str(font_path), FR_PRIVATE, 0)
    except Exception:
        # Silently ignore any platform/permission errors
        pass


def load_private_fonts(fonts: Iterable[str | Path]) -> None:
    """Load the given font files so Tkinter can use them.

    ``fonts`` is an iterable of paths to ``.ttf`` files.  Currently only
    Windows requires explicit registration; on other platforms the function
    simply verifies that the files exist.  Missing files are ignored to keep
    the loader forgiving during development.
    """

    for font in fonts:
        path = Path(font)
        if not path.exists():
            continue

        if sys.platform.startswith("win"):
            _add_windows_font(path)
        else:  # Non-Windows platforms usually pick up fonts automatically
            # Still touch the path to emphasise it is intentionally unused
            path.stat()
