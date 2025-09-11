import tkinter as tk
import os
import sys
import subprocess
from tkinter import filedialog, messagebox, ttk
from update_master import update_master

# update_master.py
def run_update(master_path, group_folder, student_cell, grade_cell, grade_column):  # CHANGED: added grade_column param
    try:
        result = update_master(master_path, group_folder, student_cell, grade_cell, grade_column)  # CHANGED: pass grade_column through
        show_success_popup(result)
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

def show_success_popup(path):
    win = tk.Toplevel()
    win.title("Success")
    win.geometry("400x100")
    tk.Label(win, text="Master file updated successfully").pack(pady=10)
    tk.Button(win, text="Open File", command=lambda: open_file_cross_platform(path)).pack()
    tk.Button(win, text="Close", command=win.destroy).pack(pady=5)

def open_file_cross_platform(path):
    try:
        if sys.platform.startswith('darwin'):
            subprocess.call(['open', path])     # macOS
        elif sys.platform.startswith('win'):
            os.startfile(path)                  # Windows
        elif sys.platform.startswith('linux'):
            subprocess.call(['xdg-open', path]) # Linux
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file:\n{e}")

# Initialize and run GUI
def init_gui():
    root = tk.Tk()
    root.title("D2L Spreadsheet Automator")
    root.geometry("500x340")  # CHANGED: slightly taller to fit new field

    # Create notebook and tab
    notebook = ttk.Notebook(root)
    tab1 = ttk.Frame(notebook)
    notebook.add(tab1, text="Grade Import")
    tab2 = ttk.Frame(notebook)
    notebook.add(tab2, text="Info")
    notebook.pack(expand=True, fill="both")

    master_var = tk.StringVar()
    group_var = tk.StringVar()
    student_cell_var = tk.StringVar(value="B4")
    grade_cell_var = tk.StringVar(value="B11")
    grade_column_var = tk.StringVar(value="G")  # ADDED: default grade column

    # Master Sheet
    tk.Label(tab1, text="Master Sheet").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(tab1, textvariable=master_var, width=50).grid(row=1, column=0, padx=10, pady=2)
    tk.Button(tab1, text="...", command=lambda: master_var.set(filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.csv")]))).grid(row=1, column=1, padx=5)

    # Group Folder
    tk.Label(tab1, text="Group Folder").grid(row=2, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(tab1, textvariable=group_var, width=50).grid(row=3, column=0, padx=10, pady=2)
    tk.Button(tab1, text="...", command=lambda: group_var.set(filedialog.askdirectory())).grid(row=3, column=1, padx=5)

    # Student Cell
    tk.Label(tab1, text="Student Names Cell").grid(row=4, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(tab1, textvariable=student_cell_var, width=6).grid(row=4, column=1, padx=5, pady=(10, 0), sticky="w")  # CHANGED: width a bit wider

    # Grade Cell
    tk.Label(tab1, text="Grade Cell").grid(row=5, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(tab1, textvariable=grade_cell_var, width=6).grid(row=5, column=1, padx=5, pady=(10, 0), sticky="w")  # CHANGED: width a bit wider

    # Grade Column (in Master)
    tk.Label(tab1, text="Master Grade Column (letter)").grid(row=6, column=0, sticky="w", padx=10, pady=(10, 0))  # ADDED
    tk.Entry(tab1, textvariable=grade_column_var, width=6).grid(row=6, column=1, padx=5, pady=(10, 0), sticky="w")  # ADDED

    # Run button
    tk.Button(
        tab1,
        text="Run Grade Import",
        command=lambda: run_update(
            master_var.get(),
            group_var.get(),
            student_cell_var.get(),
            grade_cell_var.get(),
            grade_column_var.get(),  # ADDED: pass grade column to runner
        ),
        padx=20,
        pady=10
    ).grid(row=7, column=0, columnspan=2, pady=20)  # CHANGED: row shifted down by one

    # Info tab
    info_text = tk.Text(tab2, wrap="word", width=60, height=15)
    info_text.insert("1.0", "One day, there will be information here.\n\n")
    info_text.config(state="disabled")
    info_text.pack(padx=10, pady=10, fill="both", expand=True)

    root.mainloop()

if __name__ == "__main__":
    init_gui()
