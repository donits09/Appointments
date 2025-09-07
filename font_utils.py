import os
import shutil
import sys


def _system_fonts_dir() -> str:
    """Return the OS-specific directory for installing fonts."""
    if sys.platform.startswith("win"):
        return os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts")
    if sys.platform == "darwin":
        return os.path.expanduser("~/Library/Fonts")
    return os.path.expanduser("~/.local/share/fonts")


def install_fonts(font_dir: str) -> None:
    """Copy fonts from *font_dir* to the user's system font directory if missing.

    This function is best-effort; any failure is silently ignored so that lack of
    permissions does not crash the application. It is primarily intended for
    stand-alone installers which bundle required fonts."""
    dest_dir = _system_fonts_dir()
    os.makedirs(dest_dir, exist_ok=True)
    try:
        for name in os.listdir(font_dir):
            src = os.path.join(font_dir, name)
            dst = os.path.join(dest_dir, name)
            if os.path.isfile(src) and not os.path.exists(dst):
                try:
                    shutil.copy(src, dst)
                except Exception:
                    pass
    except Exception:
        # Completely ignore any failure in font installation
        pass
