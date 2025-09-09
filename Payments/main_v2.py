import tkinter as tk
from tkinter import messagebox
from fpdf import FPDF
import csv
from datetime import datetime
import os
from tkcalendar import Calendar
import sys
from pathlib import Path

sys.path.append(str(Path(__file__).resolve().parents[1]))
from common_paths import base_dir, data_dir, find_latest_csv

# Example A: open the latest "appointments_" CSV automatically
csv_path = find_latest_csv(prefix="appointments_")
if not csv_path:
    messagebox.showerror(
        "CSV not found",
        f"No CSV found in:\n{data_dir()}\n\nUse 'Upload CSV' in the launcher first."
    )
    raise SystemExit(1)

class PDF(FPDF):
    def header(self):
        self.image(str(base_dir() / 'Header.jpg'), 27.5, 0, 150, 35)
        self.ln(20)
        self.set_font('Novecento Wide Demi Bold', 'B', 9)
        self.cell(0, 0, 'LIST OF APPOINTMENTS FOR ' + str(var_date), 'L')
        self.ln(4)
        self.set_font('Armata Regular', '', 7)
        self.cell(0, 0, 'Payments', 'L')
        self.set_font('Armata Regular', '', 7)
        self.cell(4, 0, 'Page ' + str(self.page_no()) + ' of {nb}', 0, 0, 'R')
        self.ln(2)
        self.set_font('Trebuchet MS Bold', 'B', 7)
        self.set_fill_color(180)
        self.cell(0, 5, 'No.        Res         Service                  Date                    '
                        'Start                    End                     Status                 '
                        'Client                                                    Mobile No.', 1, 1, 'L', True)
        self.ln(2)

def generate_payment_pdf():
    try:
        selected_date = cal.get_date()
        global var_date
        var_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y-%m-%d")
        c_var_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y%m%d")

        csv_filename = f"payments_{var_date}.csv"
        csv_path = data_dir() / csv_filename
        if not csv_path.exists():
            messagebox.showerror("File Not Found", f"CSV file not found for date: {var_date}")
            return

        pdf = PDF(orientation='P', unit='mm', format='A4')

        root_dir = base_dir()
        font_dir = root_dir / "Fonts"
        pdf.add_font(
            'Novecento Wide Demi Bold',
            'B',
            str(font_dir / 'Novecentowide-Bold.ttf'),
            uni=True,
        )
        pdf.add_font(
            'Armata Regular',
            '',
            str(font_dir / 'Armata-Regular.ttf'),
            uni=True,
        )
        pdf.add_font('Trebuchet MS Bold', 'B', r"C:\Windows\Fonts\trebucbd.ttf", uni=True)
        pdf.add_font('Trebuchet MS Bold Italic', 'BI', r"C:\Windows\Fonts\trebucbi.ttf", uni=True)
        pdf.add_font('Trebuchet MS Regular', '', r"C:\Windows\Fonts\trebuc.ttf", uni=True)

        pdf.add_page()
        pdf.alias_nb_pages()
        pdf.set_margins(10, 12, 12)
        page_width = pdf.w - 2 * pdf.l_margin
        th = 5

        pdf.set_font('Trebuchet MS Regular', '', 7)

        with open(csv_path, newline='', encoding='utf8') as f:
            reader = csv.reader(f)
            next(reader)
            sortedlist = filter(None, reader)
            filteredRows = sorted(sortedlist, key=lambda r_row: (
                datetime.strptime(r_row[2], "%Y-%m-%d"),
                datetime.strptime(r_row[3], "%H:%M:%S")
            ))

            ctr = 0
            for row in filteredRows:
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

        # Save inside "Payments" folder
        output_filename = f'Payment_Appointments_{c_var_date}.pdf'
        output_path = os.path.join("Payments", output_filename)
        os.makedirs("Payments", exist_ok=True)
        pdf.output(output_path, 'F')

        messagebox.showinfo("Success", f"PDF generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Tkinter GUI
root = tk.Tk()
root.withdraw()
root.title("Payment Appointments PDF Generator")
root.iconbitmap(str(base_dir() / "favicon.ico"))
root.configure(bg="white")

frame = tk.Frame(root, padx=20, pady=20, bg="white")
frame.pack()

# Date picker
tk.Label(frame, text="Select a Date:", bg="white").pack()
from tkcalendar import Calendar
cal = Calendar(frame, selectmode="day", date_pattern="mm/dd/yy")
cal.pack(pady=10)

btn_generate = tk.Button(
    frame,
    text="Generate Payment PDF",
    command=generate_payment_pdf,
    width=30,
    pady=10,
    bg="white",
    bd=1,
    relief="solid",
)
btn_generate.pack(pady=10)


def center_window(win):
    win.update_idletasks()
    width = win.winfo_reqwidth() + 20
    height = win.winfo_reqheight() + 20
    x = (win.winfo_screenwidth() - width) // 2
    y = (win.winfo_screenheight() - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


center_window(root)
root.resizable(False, False)
root.deiconify()
root.mainloop()
