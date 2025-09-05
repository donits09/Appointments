from fpdf import FPDF
import csv
from datetime import datetime
from datetime import date

now = date.today()
var_date = now.strftime("%Y-%m-%d")


class PDF(FPDF):
    def header(self):
        self.image('Header.jpg', 27.5, 0, 150, 35)
        self.ln(20)
        self.set_font('Novecento Wide Demi Bold', 'B', 9)
        self.cell(0, 0, 'LIST OF APPOINTMENTS FOR ' + str(var_date), 'L')
        self.ln(4)
        self.set_font('Armata Regular', '', 7)
        self.cell(0, 0, 'Payments', 'L')
        self.set_font('Armata Regular', '', 7)
        self.cell(4, 0, 'Page ' + str(self.page_no()) + ' of {nb}', 0, 0, 'R')
        pdf.ln(2)
        self.set_font('Trebuchet MS Bold', 'B', 7)
        self.set_fill_color(180)
        self.cell(0, 5, 'No.        Res         Service                  Date                    '
                        'Start                    End                     Status                 '
                        'Client                                                    Mobile No.', 1, 1, 'L', True)
        pdf.ln(2)


pdf = PDF(orientation='P', unit='mm', format='A4')

'''
    Change the file path of Novecento and Armata. :)    
'''

pdf.add_font('Novecento Wide Demi Bold', 'B',
             r"C:\Users\Asian Land\AppData\Local\Microsoft\Windows\Fonts\Novecentowide-Bold.ttf", uni=True)
pdf.add_font('Armata Regular', '', r"C:\Users\Asian Land\AppData\Local\Microsoft\Windows\Fonts\Armata-Regular.ttf", uni=True)
pdf.add_font('Trebuchet MS Bold', 'B', r"C:\Windows\Fonts\trebucbd.ttf", uni=True)
pdf.add_font('Trebuchet MS Bold Italic', 'BI', r"C:\Windows\Fonts\trebucbi.ttf", uni=True)
pdf.add_font('Trebuchet MS Regular', '', r"C:\Windows\Fonts\trebuc.ttf", uni=True)
pdf.add_page()
pdf.alias_nb_pages()
pdf.set_margins(10, 12, 12)
page_width = pdf.w - 2 * pdf.l_margin
col_width = page_width / 20
col_width1 = page_width / 6
col_width2 = page_width / 8
col_width3 = page_width / 10

th = 5

pdf.set_font('Trebuchet MS Regular', '', 7)


try:
    with open('payments_2025-07-10.csv', newline='', encoding='utf8') as f:
        reader = csv.reader(f)
        next(reader)
        sortedlist = filter(None, reader)
        filteredRows = sorted(sortedlist, key=lambda r_row: (datetime.strptime(r_row[2], "%Y-%m-%d"),
                                                       datetime.strptime(r_row[3], "%H:%M:%S")))
        ctr = 0
        for row in filteredRows:
            if row[2] == var_date and row[6] != "pending":
                ctr = ctr + 1
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
except:
    print("Please check your file name or the date format in your CSV file.")


pdf.set_font('Trebuchet MS Bold Italic', 'BI', 7)
pdf.ln(5)
pdf.cell(page_width, 0.0, '***Nothing Follows***', align='C')
c_var_date = now.strftime("%Y%m%d")
pdf.output(f'Payment_Appointments_{c_var_date}.pdf', 'F')
