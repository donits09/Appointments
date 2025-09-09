
# home.py
import sys, os, shutil, subprocess
from pathlib import Path

from fonts_loader import load_private_fonts  # OK at top level (no Tk side-effects)
from font_utils import install_fonts

APP_NAME = "ScriptLauncher"



def _desktop_path() -> Path | None:
    """Return the current user's Desktop directory or ``None`` if unknown."""
    try:
        from ctypes import windll, wintypes, create_unicode_buffer  # type: ignore

        CSIDL_DESKTOP = 0x10  # Desktop directory
        SHGFP_TYPE_CURRENT = 0
        buf = create_unicode_buffer(wintypes.MAX_PATH)
        if windll.shell32.SHGetFolderPathW(
            None, CSIDL_DESKTOP, None, SHGFP_TYPE_CURRENT, buf
        ) == 0:
            p = Path(buf.value)
            if p.exists():
                return p
    except Exception:
        pass

    env_path = Path(os.path.expanduser("~")) / "Desktop"
    return env_path if env_path.exists() else None


def _ps_quote(p: Path) -> str:
    """Return a single-quoted string safe for PowerShell."""
    return str(p).replace("'", "''")


def ensure_desktop_shortcut() -> None:
    """Create a shortcut to this launcher on the user's desktop (best effort)."""

    exe_path = Path(sys.executable)
    if exe_path.suffix.lower() != ".exe":
        # Only create a shortcut for a frozen EXE, skip when run with python.exe
        return


    desktop = _desktop_path()
    if not desktop:
        return

    shortcut_path = desktop / f"{APP_NAME}.lnk"
    icon = base_dir() / "favicon.ico"

    if shortcut_path.exists():

        return

    try:
        import win32com.client  # type: ignore


        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        shortcut.TargetPath = str(exe_path)
        shortcut.WorkingDirectory = str(exe_path.parent)
        if icon.exists():
            shortcut.IconLocation = str(icon)
        shortcut.Save()
        return
    except Exception:
        # fall back to PowerShell if pywin32 is unavailable or fails
        try:
            ps_parts = [
                "$ws=New-Object -ComObject WScript.Shell;",

                f"$s=$ws.CreateShortcut('{_ps_quote(shortcut_path)}');",
                f"$s.TargetPath='{_ps_quote(exe_path)}';",
                f"$s.WorkingDirectory='{_ps_quote(exe_path.parent)}';",
            ]
            if icon.exists():
                ps_parts.append(f"$s.IconLocation='{_ps_quote(icon)}';")

            ps_parts.append("$s.Save()")
            subprocess.run([
                "powershell",
                "-NoProfile",
                "-Command",
                " ".join(ps_parts),
            ])
        except Exception:
            # Best effort only – failure to create the shortcut should not abort the app
            pass

def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        # When running as a PyInstaller bundle, data files such as fonts are
        # unpacked to ``sys._MEIPASS``.  Prefer that directory so bundled
        # resources are located correctly in both one-file and one-folder
        # distributions.
        return Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
    return Path(__file__).resolve().parent


def exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def data_dir() -> Path:
    root_dir = Path(os.getenv("LOCALAPPDATA", base_dir()))
    d = root_dir / APP_NAME / "data"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _find_case_insensitive(name: str, search_dir: Path) -> Path | None:
    """Return a file within ``search_dir`` matching ``name`` regardless of casing."""
    target = name.lower()
    for p in search_dir.iterdir():
        if p.name.lower() == target:
            return p
    return None


def app_or_py(exe_name: str, fallback_rel_py: str) -> list[str]:
    """
    Prefer the sibling EXE. If missing:
      - in dev (not frozen): run the .py with the real Python
      - in frozen build: raise FileNotFoundError (caller shows a message)
    """
    exe_path = _find_case_insensitive(exe_name, exe_dir())
    if exe_path and exe_path.exists():
        return [str(exe_path)]

    if getattr(sys, "frozen", False):
        # Don't recurse by using sys.executable (that's the launcher itself)
        raise FileNotFoundError(
            f"Component '{exe_name}' was not found beside the launcher: {exe_dir() / exe_name}"
        )

    # Dev mode: run the Python script with the current interpreter
    return [sys.executable, str(base_dir() / fallback_rel_py)]

def main():
    ensure_desktop_shortcut()

    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    import tkinter.font as tkfont

    # ONE root only; hide during setup
    root = tk.Tk()
    root.withdraw()
    root.iconbitmap(str(base_dir() / "favicon.ico"))
    # Ensure relative resources resolve beside the executable
    os.chdir(base_dir())

    # Install bundled fonts (best-effort) and load them privately so Tkinter can
    # use them even on systems where they are not pre-installed.
    fonts_dir = base_dir() / "Fonts"
    install_fonts(str(fonts_dir))
    load_private_fonts([
        fonts_dir / "Armata-Regular.ttf",
        fonts_dir / "Novecentowide-Bold.ttf",
        fonts_dir / "Novecentowide-DemiBold_0.ttf",
    ])

    def choose_family(preferred_names, fallbacks=("Segoe UI", "Arial", "Tahoma")):
        fams = list(tkfont.families())
        lower = {f.lower(): f for f in fams}
        for name in list(preferred_names) + list(fallbacks):
            if name in lower:
                return lower[name]
            for f in fams:
                if name.lower() in f.lower():
                    return f
        return tkfont.nametofont("TkDefaultFont").actual("family")

    armata_family    = choose_family(["Armata"])
    novecento_family = choose_family(["Novecento Wide", "Novecentowide"])

    title_font = tkfont.Font(root=root, family=armata_family,    size=14, weight="bold")
    btn_font   = tkfont.Font(root=root, family=novecento_family, size=10)

    # ---------- UI ----------
    root.title("Script Launcher")
    width, height = 550, 420
    root.geometry(f"{width}x{height}")
    root.update_idletasks()
    x = (root.winfo_screenwidth() - width) // 2
    y = (root.winfo_screenheight() - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.resizable(False, False)
    root.configure(bg="#ffffff")

    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure("TFrame", background="#ffffff")
    style.configure(
        "TButton",
        background="#ffffff",
        padding=6,
        relief="solid",
        borderwidth=1,
    )
    style.map("TButton", background=[("active", "#d4d4d4")])
    style.configure("TLabel", background="#ffffff")
    style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", bordercolor="#d9d9d9")
    style.configure("Treeview.Heading", font=btn_font, background="#d9d9d9")

    button_frame = ttk.Frame(root, padding=20)
    button_frame.pack(side="left", fill="y", padx=10)

    treeview_frame = ttk.Frame(root, padding=20)
    treeview_frame.pack(side="right", fill="y", padx=10)

    ttk.Label(button_frame, text="Appointments Scripts", font=title_font).pack(pady=(0, 15))

    ttk.Label(treeview_frame, text="CSV Files (User Data):", font=btn_font).pack(pady=(10, 5))

    treeview = ttk.Treeview(treeview_frame, columns=("File Name",), show="headings", height=12, selectmode="extended")
    treeview.heading("File Name", text="File Name")
    treeview.pack(side="top", fill="both", expand=True, pady=5)

    button_width = 20

    # ---- handlers ----
    def refresh_csv_list():
        csv_files = [f.name for f in data_dir().glob("*.csv")]
        for row in treeview.get_children():
            treeview.delete(row)
        for file in csv_files:
            treeview.insert("", "end", values=(file,))

    def upload_csv_file():
        file_paths = filedialog.askopenfilenames(
            filetypes=[("CSV files", "*.csv")],
            title="Select CSV file(s)"
        )
        if not file_paths:
            return
        try:
            dst = data_dir()
            for file_path in file_paths:
                filename = os.path.basename(file_path)
                shutil.copy(file_path, dst / filename)
            messagebox.showinfo("Upload Complete", f"Successfully uploaded {len(file_paths)} file(s).")
            refresh_csv_list()
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def delete_csv_files():
        selected_items = treeview.selection()
        if not selected_items:
            messagebox.showwarning("Selection Error", "Please select files to delete.")
            return
        if not messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected files?"):
            return
        for selected_item in selected_items:
            file_name = treeview.item(selected_item, 'values')[0]
            file_path = data_dir() / file_name
            try:
                if file_path.exists():
                    file_path.unlink()
                treeview.delete(selected_item)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to delete {file_name}:\n{e}")
        messagebox.showinfo("Delete Complete", f"Successfully deleted {len(selected_items)} file(s).")

    def run_component(exe_name, fallback_rel_py, title):
        try:
            cmd = app_or_py(exe_name, fallback_rel_py)
            subprocess.Popen(cmd, shell=False)
        except FileNotFoundError as miss:
            message = (
                f'{miss}\n\n'
                'Copy all component EXEs beside ScriptLauncher.exe, or reinstall:\n'
                f' - {exe_name}'
            )
            messagebox.showerror('Component missing', message)
        except Exception as e:
            messagebox.showerror('Error', f'Failed to run {title}:\n{e}')

    def run_appointments():
        run_component("Appointments.exe", "Appointments/main_v2.py", "Appointments")

    def run_payments():
        run_component("Payments.exe", "Payments/main_v2.py", "Payments")

    def run_pending():
        run_component("Pending.exe", "Pending/main_v2.py", "Pending")

    def run_pdf_viewer():
        run_component("PDFViewer.exe", "pdf.py", "PDF Viewer")

    # ---- buttons ----
    ttk.Button(button_frame, text="Upload CSV", command=upload_csv_file, width=button_width).pack(pady=5, ipady=5)
    ttk.Button(button_frame, text="Appointments", command=run_appointments, width=button_width).pack(pady=5, ipady=5)
    ttk.Button(button_frame, text="Payments", command=run_payments, width=button_width).pack(pady=5, ipady=5)
    ttk.Button(button_frame, text="Pending", command=run_pending, width=button_width).pack(pady=5, ipady=5)
    ttk.Button(button_frame, text="View PDF", command=run_pdf_viewer, width=button_width).pack(pady=5, ipady=5)
    ttk.Button(button_frame, text="Exit", command=root.quit, width=button_width).pack(pady=(10, 5), ipady=5)

    ttk.Button(treeview_frame, text="Delete File", command=delete_csv_files, width=button_width).pack(side="bottom", pady=10)

    # initial load
    refresh_csv_list()

    # Optional: warn early if some components are missing (in frozen build)
    if getattr(sys, "frozen", False):
        existing = {p.name.lower() for p in base_dir().iterdir() if p.is_file()}
        expected = ("Appointments.exe", "Payments.exe", "Pending.exe", "PDFViewer.exe")
        present = [n for n in expected if n.lower() in existing]
        if present and len(present) < len(expected):
            missing = [n for n in expected if n.lower() not in existing]
            messagebox.showwarning(
                "Missing components",
                "These EXEs are not beside ScriptLauncher.exe:\n- " + "\n- ".join(missing)
            )

    # Show window now that setup is complete
    root.deiconify()
    root.mainloop()

if __name__ == "__main__":
    main()

