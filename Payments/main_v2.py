import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from fpdf import FPDF
import csv
from datetime import datetime, date
import os
import shutil
import sys, os

def resource_path(*parts) -> str:
    """Returns absolute path to a bundled resource (PyInstaller-compatible)."""
    base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, *parts)

def app_dir() -> str:
    """Where we read/write CSVs/PDFs."""
    if getattr(sys, "_MEIPASS", None):
        # Next to the EXE (Program Files may be read-only; see Note below)
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

script_dir = app_dir()

ASSETS_HEADER = resource_path("Header.jpg")
FONTS_DIR = resource_path("Fonts")


# -------------------- Globals --------------------
HEADER_TITLE = ""           # e.g., "LIST OF APPOINTMENTS FOR " or "LIST OF PENDING APPOINTMENTS AS OF "
HEADER_SUBTITLE = ""        # e.g., "ALSC Services" or "Payments"
HEADER_COLUMNS_TEXT = ""    # table header row text
var_date = ""               # date string like "2025-09-03" set per run

script_dir = os.path.dirname(os.path.abspath(__file__))


# -------------------- PDF Class --------------------
class PDF(FPDF):
    def header(self):

        try:
            if os.path.exists(ASSETS_HEADER):
                self.image(ASSETS_HEADER, 27.5, 0, 150, 35)
        except Exception:
            pass

        self.ln(20)
        self.set_font('Novecento Wide Demi Bold', 'B', 9)
        title_text = (HEADER_TITLE or "LIST OF APPOINTMENTS FOR ") + str(var_date)
        self.cell(0, 0, title_text, 'L')
        self.ln(4)

        self.set_font('Armata Regular', '', 7)
        self.cell(0, 0, HEADER_SUBTITLE, 'L')

        self.set_font('Armata Regular', '', 7)
        self.cell(4, 0, 'Page ' + str(self.page_no()) + ' of {nb}', 0, 0, 'R')
        self.ln(2)

        # Table header
        self.set_font('Trebuchet MS Bold', 'B', 7)
        self.set_fill_color(180)
        if HEADER_COLUMNS_TEXT:
            self.multi_cell(0, 5, HEADER_COLUMNS_TEXT, 1, 'L', True)
        else:
            self.cell(0, 5, 'No.  ...', 1, 1, 'L', True)
        self.ln(2)


def register_fonts(pdf):
    try:
        pdf.add_font('Novecento Wide Demi Bold', 'B', os.path.join(FONTS_DIR, 'Novecentowide-Bold.ttf'), uni=True)
    except Exception: pass
    try:
        pdf.add_font('Armata Regular', '', os.path.join(FONTS_DIR, 'Armata-Regular.ttf'), uni=True)
    except Exception: pass
    for fam, style, fname in [
        ('Trebuchet MS Bold', 'B', 'trebucbd.ttf'),
        ('Trebuchet MS Bold Italic', 'BI', 'trebucbi.ttf'),
        ('Trebuchet MS Regular', '', 'trebuc.ttf'),
    ]:
        try:
            pdf.add_font(fam, style, os.path.join(FONTS_DIR, fname), uni=True)
        except Exception:
            pass


def generate_all_pdfs():
    """Generate Payments + Pending + Appointments PDFs for TODAY (ignores the entry)."""
    # Force today's date (mm/dd/yy) into the entry so all generators use it
    today_str = date.today().strftime("%m/%d/%y")
    date_entry.delete(0, tk.END)
    date_entry.insert(0, today_str)

    # Validate the date (should always pass)
    try:
        parse_and_validate_date(today_str)
    except ValueError as e:
        messagebox.showerror("Error", str(e))
        return

    # Suppress individual success popups; let errors show as usual
    orig_showinfo = messagebox.showinfo
    def _silent_info(*args, **kwargs):
        return None
    messagebox.showinfo = _silent_info

    finished = []
    try:
        generate_payment_pdf()
        finished.append("Payments")

        generate_pending_pdf()
        finished.append("Pending")

        generate_appointment_pdf()
        finished.append("Appointments")
    finally:
        messagebox.showinfo = orig_showinfo

    if finished:
        messagebox.showinfo("Generate All (Today)", f"Finished generating for {today_str}: {', '.join(finished)}")
    else:
        messagebox.showwarning("Generate All (Today)", "No PDFs were generated. Check for errors above.")

def parse_and_validate_date(mmddyy: str):
    """Parses mm/dd/yy to (var_date 'YYYY-MM-DD', c_var_date 'YYYYMMDD', month 'MM')."""
    try:
        d = datetime.strptime(mmddyy.strip(), "%m/%d/%y")
    except ValueError:
        raise ValueError("Please enter a valid date in mm/dd/yy format.")
    return d.strftime("%Y-%m-%d"), d.strftime("%Y%m%d"), d.strftime("%m")


# -------------------- GENERATORS --------------------
def generate_payment_pdf():
    global var_date, HEADER_TITLE, HEADER_SUBTITLE, HEADER_COLUMNS_TEXT

    try:
        selected_date = date_entry.get().strip()
        if not selected_date:
            messagebox.showerror("Error", "Please enter a date (mm/dd/yy).")
            return

        var_date, c_var_date, _ = parse_and_validate_date(selected_date)

        csv_filename = os.path.join(script_dir, f'payments_{var_date}.csv')
        if not os.path.exists(csv_filename):
            messagebox.showerror("File Not Found", f"CSV file not found for date: {var_date}")
            return

        # Header configuration
        HEADER_TITLE = "LIST OF APPOINTMENTS FOR "
        HEADER_SUBTITLE = "Payments"
        HEADER_COLUMNS_TEXT = (
            'No.        Res         Service                  Date                    '
            'Start                    End                     Status                 '
            'Client                                                    Mobile No.'
        )

        pdf = PDF(orientation='P', unit='mm', format='A4')
        register_fonts(pdf)
        pdf.add_page()
        pdf.alias_nb_pages()
        pdf.set_margins(10, 12, 12)
        page_width = pdf.w - 2 * pdf.l_margin
        th = 5
        pdf.set_font('Trebuchet MS Regular', '', 7)

        with open(csv_filename, newline='', encoding='utf8') as f:
            reader = csv.reader(f)
            next(reader, None)
            sortedlist = filter(None, reader)
            filteredRows = sorted(
                sortedlist,
                key=lambda r_row: (
                    datetime.strptime(r_row[2], "%Y-%m-%d"),   # Date
                    datetime.strptime(r_row[3], "%H:%M:%S")    # Start
                )
            )

            ctr = 0
            for row in filteredRows:
                # 0:Res, 1:Service, 2:Date, 3:Start, 4:End, 6:Status, 8:Client, 9:Mobile
                if row[2] == var_date and row[6].lower() != "pending":
                    ctr += 1
                    pdf.set_fill_color(180)
                    pdf.cell(10, th, str(ctr))
                    pdf.cell(10, th, str(row[0]))
                    pdf.cell(20, th, str(row[1]))
                    pdf.cell(20, th, str(row[2]))
                    pdf.cell(20, th, str(row[3]))
                    pdf.cell(20, th, str(row[4]))
                    pdf.cell(20, th, str(row[6]))
                    pdf.cell(44, th, str(row[8]))
                    pdf.cell(20, th, str(row[9]))
                    pdf.ln(1)
                    pdf.ln(5)

        pdf.set_font('Trebuchet MS Bold Italic', 'BI', 7)
        pdf.ln(5)
        pdf.cell(page_width, 0.0, '***Nothing Follows***', align='C')

        # Save
        out_dir = os.path.join(script_dir, "Payments")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f'Payment_Appointments_{c_var_date}.pdf')
        pdf.output(output_path, 'F')

        messagebox.showinfo("Success", f"PDF generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def generate_appointment_pdf():
    global var_date, HEADER_TITLE, HEADER_SUBTITLE, HEADER_COLUMNS_TEXT

    try:
        selected_date = date_entry.get().strip()
        if not selected_date:
            messagebox.showerror("Error", "Please enter a date (mm/dd/yy).")
            return

        var_date, c_var_date, _ = parse_and_validate_date(selected_date)

        csv_filename = os.path.join(script_dir, f'appointments_{var_date}.csv')
        if not os.path.exists(csv_filename):
            messagebox.showerror("File Not Found", f"CSV file not found for date: {var_date}")
            return

        # Header configuration
        HEADER_TITLE = "LIST OF APPOINTMENTS FOR "
        HEADER_SUBTITLE = "ALSC Services"
        HEADER_COLUMNS_TEXT = (
            'No.        Res        Service                                        '
            'Start             End              Status               '
            'Client                                                                   '
            'Contract Supp.'
        )

        pdf = PDF(orientation='P', unit='mm', format='A4')
        register_fonts(pdf)
        pdf.add_page()
        pdf.alias_nb_pages()
        pdf.set_margins(10, 12, 12)
        page_width = pdf.w - 2 * pdf.l_margin
        th = 5
        pdf.set_font('Trebuchet MS Regular', '', 7)

        with open(csv_filename, 'r', newline='', encoding='utf8') as f:
            reader = csv.reader(f)
            next(reader, None)
            sortedlist = filter(None, reader)
            filteredRows = sorted(
                sortedlist,
                key=lambda r_row: (
                    datetime.strptime(r_row[4], "%Y-%m-%d"),  # Date
                    r_row[8],                                # Status (tie-breaker)
                    datetime.strptime(r_row[6], "%H:%M:%S")  # End time (as in your script)
                )
            )

            ctr = 0
            for row in filteredRows:
                # 0:Res, 3:Service, 4:Date, 5:Start, 6:End, 8:Status, 9:Client, 21:Contract Supp.
                if row[4] == var_date:
                    ctr += 1
                    pdf.set_fill_color(180)
                    pdf.cell(10, th, str(ctr))
                    pdf.cell(10, th, str(row[0]))
                    pdf.cell(35, th, str(row[3]))
                    pdf.cell(15, th, str(row[5]))
                    pdf.cell(15, th, str(row[6]))
                    pdf.cell(18, th, str(row[8]))
                    pdf.cell(35, th, str(row[9]))
                    pdf.cell(30, th, str(row[21]))
                    pdf.ln(1)
                    pdf.ln(5)

        pdf.set_font('Trebuchet MS Bold Italic', 'BI', 7)
        pdf.ln(5)
        pdf.cell(page_width, 0.0, '***Nothing Follows***', align='C')

        # Save
        out_dir = os.path.join(script_dir, "Appointments")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f'Appointments_{c_var_date}.pdf')
        pdf.output(output_path, 'F')

        messagebox.showinfo("Success", f"PDF generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def generate_pending_pdf():
    global var_date, HEADER_TITLE, HEADER_SUBTITLE, HEADER_COLUMNS_TEXT

    try:
        selected_date = date_entry.get().strip()
        if not selected_date:
            messagebox.showerror("Error", "Please enter a date (mm/dd/yy).")
            return

        var_date, c_var_date, var_date_num = parse_and_validate_date(selected_date)  # var_date_num = "MM"

        csv_filename = os.path.join(script_dir, f'appointments_{var_date}.csv')
        if not os.path.exists(csv_filename):
            messagebox.showerror("File Not Found", f"CSV file not found: {csv_filename}")
            return

        # Header configuration
        HEADER_TITLE = "LIST OF PENDING APPOINTMENTS AS OF "
        HEADER_SUBTITLE = "ALSC Services"
        HEADER_COLUMNS_TEXT = (
            'No.   Res        Service                                 Date                '
            'Start            End              Status            '
            'Client                                     '
            'Contract Supp.            Mobile No.'
        )

        pdf = PDF(orientation='P', unit='mm', format='A4')
        register_fonts(pdf)
        pdf.add_page()
        pdf.alias_nb_pages()
        pdf.set_margins(10, 12, 12)
        page_width = pdf.w - 2 * pdf.l_margin
        th = 5
        pdf.set_font('Trebuchet MS Regular', '', 7)

        with open(csv_filename, newline='', encoding='utf8') as f:
            reader = csv.reader(f)
            next(reader, None)
            sortedlist = filter(None, reader)
            filteredRows = sorted(
                sortedlist,
                key=lambda r_row: (
                    datetime.strptime(r_row[4], "%Y-%m-%d"),  # Date
                    datetime.strptime(r_row[5], "%H:%M:%S"),  # Start
                    datetime.strptime(r_row[6], "%H:%M:%S")   # End
                )
            )

            ctr = 0
            for row in filteredRows:
                # Filter: month >= selected month AND status == pending
                appointment_month = row[4][5:7]  # MM from YYYY-MM-DD
                if appointment_month >= var_date_num and row[8].lower() == "pending":
                    ctr += 1
                    pdf.set_fill_color(180)
                    pdf.cell(6, th, str(ctr))
                    pdf.cell(10, th, str(row[0]))     # Res
                    pdf.cell(30, th, str(row[3]))     # Service
                    pdf.cell(17, th, str(row[4]))     # Date
                    pdf.cell(15, th, str(row[5]))     # Start
                    pdf.cell(15, th, str(row[6]))     # End
                    pdf.cell(15, th, str(row[8]))     # Status
                    pdf.cell(35, th, str(row[9]))     # Client
                    pdf.cell(26, th, str(row[21]))    # Contract Supp.
                    pdf.cell(30, th, str(row[11]))    # Mobile No.
                    pdf.ln(1)
                    pdf.ln(5)

        pdf.set_font('Trebuchet MS Bold Italic', 'BI', 7)
        pdf.ln(5)
        pdf.cell(page_width, 0.0, '***Nothing Follows***', align='C')

        # Save
        out_dir = os.path.join(script_dir, "Pending")
        os.makedirs(out_dir, exist_ok=True)
        output_path = os.path.join(out_dir, f'Pending_Appointments_{c_var_date}.pdf')
        pdf.output(output_path, 'F')

        messagebox.showinfo("Success", f"PDF generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# -------------------- CSV Utilities (moved here) --------------------
def upload_csv_file():
    file_paths = filedialog.askopenfilenames(
        filetypes=[("CSV files", "*.csv")],
        title="Select CSV file(s)"
    )
    if not file_paths:
        return

    try:
        for file_path in file_paths:
            filename = os.path.basename(file_path)
            destination_path = os.path.join(script_dir, filename)
            shutil.copy(file_path, destination_path)

        messagebox.showinfo("Upload Complete", f"Successfully uploaded {len(file_paths)} file(s).")
        refresh_csv_list()
    except Exception as e:
        messagebox.showerror("Error", str(e))


def refresh_csv_list():
    """Refresh the list of CSV files in the treeview."""
    csv_files = [f for f in os.listdir(script_dir) if f.lower().endswith('.csv')]
    for row in treeview.get_children():
        treeview.delete(row)
    for file in csv_files:
        treeview.insert("", "end", values=(file,))


def delete_csv_files():
    """Delete selected CSV files from directory and treeview."""
    selected_items = treeview.selection()
    if not selected_items:
        messagebox.showwarning("Selection Error", "Please select files to delete.")
        return

    confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected files?")
    if not confirm:
        return

    errors = []
    for selected_item in selected_items:
        file_name = treeview.item(selected_item, 'values')[0]
        file_path = os.path.join(script_dir, file_name)
        try:
            os.remove(file_path)
            treeview.delete(selected_item)
        except Exception as e:
            errors.append(f"{file_name}: {e}")

    if errors:
        messagebox.showerror("Some files failed to delete", "\n".join(errors))
    else:
        messagebox.showinfo("Delete Complete", "Selected file(s) deleted.")
    refresh_csv_list()


def view_pdf():
    """Open a PDF file using the default viewer."""
    # Prefer the output folders; start in script_dir
    file_path = filedialog.askopenfilename(
        initialdir=script_dir,
        title="Open PDF",
        filetypes=[("PDF files", "*.pdf")]
    )
    if not file_path:
        return
    try:
        os.startfile(file_path)  # Windows
    except AttributeError:
        # Cross-platform fallback
        import subprocess, sys
        if sys.platform.startswith('darwin'):
            subprocess.call(('open', file_path))
        else:
            subprocess.call(('xdg-open', file_path))


# -------------------- UI --------------------
root = tk.Tk()
root.title("Appointments / Payments / Pending PDF Generator")
root.geometry("820x460")
root.resizable(False, False)

style = ttk.Style(root)
try:
    style.theme_use("clam")
except Exception:
    pass

# Left: controls
left = ttk.Frame(root, padding=16)
left.pack(side="left", fill="y")

# Right: CSV list
right = ttk.Frame(root, padding=16)
right.pack(side="right", fill="both", expand=True)

ttk.Label(left, text="Enter Date (mm/dd/yy):").pack(anchor="w")
date_entry = ttk.Entry(left, width=20)
date_entry.insert(0, date.today().strftime("%m/%d/%y"))
date_entry.pack(pady=(0, 12), anchor="w")

btn_w = 26
ttk.Button(left, text="Generate ALL (Appts + Payments + Pending)",
           command=generate_all_pdfs, width=btn_w).pack(pady=(0, 8))

ttk.Button(left, text="Generate Appointments PDF", command=generate_appointment_pdf, width=btn_w).pack(pady=4)
ttk.Button(left, text="Generate Payments PDF", command=generate_payment_pdf, width=btn_w).pack(pady=4)
ttk.Button(left, text="Generate Pending PDF", command=generate_pending_pdf, width=btn_w).pack(pady=4)

ttk.Separator(left, orient="horizontal").pack(fill="x", pady=8)

ttk.Button(left, text="Upload CSV", command=upload_csv_file, width=btn_w).pack(pady=4)
ttk.Button(left, text="View PDF", command=view_pdf, width=btn_w).pack(pady=4)
ttk.Button(left, text="Exit", command=root.quit, width=btn_w).pack(pady=(12, 4))

# Right panel: CSV files tree
ttk.Label(right, text="CSV Files in Directory:").pack(anchor="w")
treeview = ttk.Treeview(right, columns=("File Name",), show="headings", height=16, selectmode="extended")
treeview.heading("File Name", text="File Name")
treeview.column("File Name", width=420)
treeview.pack(fill="both", expand=True, pady=(6, 6))

btns_frame = ttk.Frame(right)
btns_frame.pack(fill="x")
ttk.Button(btns_frame, text="Delete File(s)", command=delete_csv_files, width=20).pack(side="left", padx=(0, 8))
ttk.Button(btns_frame, text="Refresh List", command=refresh_csv_list, width=20).pack(side="left")

# Initial list
refresh_csv_list()

root.mainloop()