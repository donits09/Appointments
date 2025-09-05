import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as font
import os
import shutil

def upload_csv_file():
    file_paths = filedialog.askopenfilenames(
        filetypes=[("CSV files", "*.csv")],
        title="Select CSV file(s)"
    )
    if not file_paths:
        return 

    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        for file_path in file_paths:
            filename = os.path.basename(file_path)
            destination_path = os.path.join(script_dir, filename)
            shutil.copy(file_path, destination_path)

        messagebox.showinfo("Upload Complete", f"Successfully uploaded {len(file_paths)} file(s).")
        refresh_csv_list()
    except Exception as e:
        messagebox.showerror("Error", str(e))

def refresh_csv_list():
    """ Refresh the list of CSV files in the treeview. """
    csv_files = [f for f in os.listdir(script_dir) if f.endswith('.csv')]
    for row in treeview.get_children():
        treeview.delete(row)
    for file in csv_files:
        treeview.insert("", "end", values=(file,))

def delete_csv_files():
    """ Delete selected CSV files from the directory and treeview. """
    selected_items = treeview.selection()
    if not selected_items:
        messagebox.showwarning("Selection Error", "Please select files to delete.")
        return

    confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected files?")
    if not confirm:
        return

    for selected_item in selected_items:
        file_name = treeview.item(selected_item, 'values')[0]
        file_path = os.path.join(script_dir, file_name)

        try:
            os.remove(file_path)
            treeview.delete(selected_item)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete {file_name}:\n{e}")

    messagebox.showinfo("Delete Complete", f"Successfully deleted {len(selected_items)} file(s).")

def run_tkinter_script():
    try:
        subprocess.Popen(["python", "Appointments/main_v2.py"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run Appointments script:\n{e}")

def run_tkinter2_script():
    try:
        subprocess.Popen(["python", "Payments/main_v2.py"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run Payments script:\n{e}")

def run_tkinter3_script():
    try:
        subprocess.Popen(["python", "Pending/main_v2.py"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run Pending script:\n{e}")

def run_tkinter4_script():
    try:
        subprocess.Popen(["python", "pdf.py"], shell=True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to run PDF script:\n{e}")

# Main window
root = tk.Tk()
root.title("Script Launcher")
root.geometry("550x420")
root.resizable(False, False)
root.configure(bg="#f0f0f0")

style = ttk.Style(root)
style.theme_use("clam")
style.configure("TFrame", background="#f0f0f0")
style.configure("TButton", background="#f0f0f0", padding=6)
style.configure("TLabel", background="#f0f0f0")

title_font = font.Font(family="Segoe UI", size=14, weight="bold")
btn_font = font.Font(family="Segoe UI", size=10)

button_frame = ttk.Frame(root, padding=20)
button_frame.pack(side="left", fill="y", padx=10)

treeview_frame = ttk.Frame(root, padding=20)
treeview_frame.pack(side="right", fill="y", padx=10)

title_label = ttk.Label(button_frame, text="Appointments Scripts", font=title_font)
title_label.pack(pady=(0, 15))

button_width = 20

ttk.Button(button_frame, text="Upload CSV", command=upload_csv_file, width=button_width).pack(pady=5, ipady=5)
ttk.Button(button_frame, text="Appointments", command=run_tkinter_script, width=button_width).pack(pady=5, ipady=5)
ttk.Button(button_frame, text="Payments", command=run_tkinter2_script, width=button_width).pack(pady=5, ipady=5)
ttk.Button(button_frame, text="Pending", command=run_tkinter3_script, width=button_width).pack(pady=5, ipady=5)
ttk.Button(button_frame, text="View PDF", command=run_tkinter4_script, width=button_width).pack(pady=5, ipady=5)
ttk.Button(button_frame, text="Exit", command=root.quit, width=button_width).pack(pady=(10, 5), ipady=5)

script_dir = os.path.dirname(os.path.abspath(__file__))
csv_files = [f for f in os.listdir(script_dir) if f.endswith('.csv')]

treeview_label = ttk.Label(treeview_frame, text="CSV Files in Directory:")
treeview_label.pack(pady=(10, 5))

treeview = ttk.Treeview(treeview_frame, columns=("File Name",), show="headings", height=12, selectmode="extended")
treeview.heading("File Name", text="File Name")
treeview.pack(side="top", fill="both", expand=True, pady=5)

for file in csv_files:
    treeview.insert("", "end", values=(file,))

ttk.Button(treeview_frame, text="Delete File", command=delete_csv_files, width=button_width).pack(side="bottom", pady=10)

root.mainloop()
