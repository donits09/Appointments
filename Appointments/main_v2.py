import tkinter as tk
from tkinter import messagebox
from fpdf import FPDF
import csv
from datetime import datetime, date
import os
from tkcalendar import Calendar
from common_paths import data_dir, find_latest_csv

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
        self.image('Header.jpg', 27.5, 0, 150, 35)
        self.ln(20)
        self.set_font('Novecento Wide Demi Bold', 'B', 9)
        self.cell(0, 0, 'LIST OF APPOINTMENTS FOR ' + str(var_date), 'L')
        self.ln(4)
        self.set_font('Armata Regular', '', 7)
        self.cell(0, 0, 'ALSC Services', 'L')
        self.set_font('Armata Regular', '', 7)
        self.cell(4, 0, 'Page ' + str(self.page_no()) + ' of {nb}', 0, 0, 'R')
        self.ln(2)
        self.set_font('Trebuchet MS Bold', 'B', 7)
        self.set_fill_color(180)
        self.cell(0, 5, 'No.        Res        Service                                        '
                        'Start             End              Status               '
                        'Client                                                                   '
                        'Contract Supp.', 1, 1, 'L', fill=True)
        self.ln(2)

def generate_pdf():
    try:
        selected_date = cal.get_date()  # Get the date selected by the user
        global var_date
        var_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y-%m-%d")
        c_var_date = datetime.strptime(selected_date, "%m/%d/%y").strftime("%Y%m%d")

        csv_filename = f'appointments_{var_date}.csv'
        
        if not os.path.exists(csv_filename):
            messagebox.showerror("File Not Found", f"CSV file not found for date: {var_date}")
            return

        pdf = PDF(orientation='P', unit='mm', format='A4')
        pdf.add_font('Novecento Wide Demi Bold', 'B',
                     r"C:\Users\donit\AppData\Local\Microsoft\Windows\Fonts\Novecentowide-Bold.ttf", uni=True)
        pdf.add_font('Armata Regular', '', r"C:\Users\donit\AppData\Local\Microsoft\Windows\Fonts\Armata-Regular.ttf", uni=True)
        pdf.add_font('Trebuchet MS Bold', 'B', r"C:\Windows\Fonts\trebucbd.ttf", uni=True)
        pdf.add_font('Trebuchet MS Bold Italic', 'BI', r"C:\Windows\Fonts\trebucbi.ttf", uni=True)
        pdf.add_font('Trebuchet MS Regular', '', r"C:\Windows\Fonts\trebuc.ttf", uni=True)

        pdf.add_page()
        pdf.alias_nb_pages()
        pdf.set_margins(10, 12, 12)
        page_width = pdf.w - 2 * pdf.l_margin
        th = 5

        pdf.set_font('Trebuchet MS Regular', '', 7)

        with open(csv_filename, 'r', newline='', encoding='utf8') as f:
            reader = csv.reader(f)
            next(reader)
            sortedlist = filter(None, reader)
            filteredRows = sorted(sortedlist, key=lambda r_row: (datetime.strptime(r_row[4], "%Y-%m-%d"),
                                                                 r_row[8],
                                                                 datetime.strptime(r_row[6], "%H:%M:%S")))

            ctr = 0
            for row in filteredRows:
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

        # Save inside "Appointments" folder
        output_filename = f'Appointments_{c_var_date}.pdf'
        output_path = os.path.join("Appointments", output_filename)
        os.makedirs("Appointments", exist_ok=True)
        pdf.output(output_path, 'F')

        messagebox.showinfo("Success", f"PDF generated:\n{output_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# Tkinter GUI
root = tk.Tk()
root.title("Appointment PDF Generator")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

# Date picker (Calendar widget)
cal_label = tk.Label(frame, text="Select a Date:")
cal_label.pack()

cal = Calendar(frame, selectmode="day", date_pattern="mm/dd/yy")
cal.pack(pady=10)

btn_generate = tk.Button(frame, text="Load & Generate PDF", command=generate_pdf, width=30, pady=10)
btn_generate.pack(pady=10)

root.mainloop()
