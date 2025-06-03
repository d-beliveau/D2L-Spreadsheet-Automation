import tkinter as tk
import os
import sys
import subprocess
from tkinter import filedialog, messagebox
from update_master import update_master

# update_master.py
def run_update(master_path, group_folder):
    try:
        result = update_master(master_path, group_folder)
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
    root.title("Grade Import Tool")
    root.geometry("500x300")

    # Master sheet
    master_var = tk.StringVar()
    tk.Label(root, text="Master Sheet").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(root, textvariable=master_var, width=60).grid(row=1, column=0, padx=10)
    tk.Button(root, text="...", command=lambda: master_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).grid(row=1, column=1, padx=5)

    # Group folder
    group_var = tk.StringVar()
    tk.Label(root, text="Group Folder").grid(row=2, column=0, sticky="w", padx=10, pady=(10, 0))
    tk.Entry(root, textvariable=group_var, width=60).grid(row=3, column=0, padx=10)
    tk.Button(root, text="...", command=lambda: group_var.set(filedialog.askdirectory())).grid(row=3, column=1, padx=5)

    # Run button
    tk.Button(root, text="Run Grade Import", command=lambda: run_update(master_var.get(), group_var.get()), padx=20, pady=10).grid(row=4, column=0, pady=20, sticky="w", padx=10)

    root.mainloop()

if __name__ == "__main__":
    init_gui()
