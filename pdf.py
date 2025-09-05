import tkinter as tk
from tkinter import ttk, messagebox
import os
import glob
import platform
import subprocess

def open_pdf(filepath):
    try:
        if platform.system() == "Windows":
            os.startfile(filepath)
        elif platform.system() == "Darwin":
            subprocess.run(["open", filepath])
        else:
            subprocess.run(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("Error", f"Could not open PDF:\n{e}")

def show_pdf_list(folder_name):
    os.makedirs(folder_name, exist_ok=True)
    pdf_files = sorted(glob.glob(os.path.join(folder_name, "*.pdf")), key=os.path.getctime, reverse=True)

    if not pdf_files:
        messagebox.showinfo("No PDFs Found", f"No PDF files found in the '{folder_name}' folder.")
        return

    def open_selected_pdf():
        selected = pdf_combo.get()
        if selected:
            filepath = os.path.join(folder_name, selected)
            window.destroy()
            open_pdf(filepath)

    def delete_selected_pdf():
        selected = pdf_combo.get()
        if selected:
            filepath = os.path.join(folder_name, selected)
            confirm = messagebox.askyesno("Delete PDF", f"Are you sure you want to delete '{selected}'?")
            if confirm:
                try:
                    os.remove(filepath)
                    messagebox.showinfo("Deleted", f"'{selected}' has been deleted.")
                    # Refresh the combo box
                    updated_files = sorted(glob.glob(os.path.join(folder_name, "*.pdf")), key=os.path.getctime, reverse=True)
                    pdf_combo['values'] = [os.path.basename(f) for f in updated_files]
                    if updated_files:
                        pdf_combo.set(os.path.basename(updated_files[0]))
                    else:
                        window.destroy()
                        messagebox.showinfo("No PDFs Left", f"All PDF files in '{folder_name}' have been deleted.")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not delete PDF:\n{e}")

    window = tk.Toplevel()
    window.title(f"{folder_name} PDFs")
    window.geometry("400x220")
    window.resizable(False, False)

    ttk.Label(window, text=f"Select a PDF from {folder_name}", font=("Segoe UI", 11)).pack(pady=10)

    pdf_combo = ttk.Combobox(window, values=[os.path.basename(f) for f in pdf_files], state="readonly", width=50)
    pdf_combo.set(os.path.basename(pdf_files[0]))
    pdf_combo.pack(pady=5)

    ttk.Button(window, text="📂 Open", command=open_selected_pdf).pack(pady=5, ipady=5)
    ttk.Button(window, text="🗑️ Delete", command=delete_selected_pdf).pack(pady=5, ipady=5)

# Main Window
root = tk.Tk()
root.title("Open a PDF File")
root.geometry("360x250")
root.resizable(False, False)

ttk.Label(root, text="Choose a PDF to view:", font=("Segoe UI", 13, "bold")).pack(pady=15)

ttk.Button(root, text="📄 Appointments", command=lambda: show_pdf_list("Appointments"), width=30).pack(pady=5, ipady=5)
ttk.Button(root, text="💰 Payments", command=lambda: show_pdf_list("Payments"), width=30).pack(pady=5, ipady=5)
ttk.Button(root, text="🕒 Pending", command=lambda: show_pdf_list("Pending"), width=30).pack(pady=5, ipady=5)

root.mainloop()
