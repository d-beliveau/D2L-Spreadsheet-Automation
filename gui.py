import tkinter as tk
from tkinter import filedialog, messagebox
from update_master import update_master

def run_update():
    master_path = filedialog.askopenfilename(
        title="Select Master Sheet",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if not master_path:
        return
    master_var.set(master_path)

    working_dir = filedialog.askdirectory(title="Select Working Directory")
    if not working_dir:
        return
    working_dir_var.set(working_dir)

    group_folder = filedialog.askdirectory(title="Select Group Folder")
    if not group_folder:
        return
    group_var.set(group_folder)

    try:
        result = update_master(master_path, group_folder)
        messagebox.showinfo("Success", f"Master file updated:\n{result}")
    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

if __name__ == "__main__":
    # Set up the GUI
    root = tk.Tk()
    root.title("Grade Import Tool")
    root.geometry("500x300")

    master_var = tk.StringVar()
    group_var = tk.StringVar()
    working_dir_var = tk.StringVar()

    tk.Label(root, text="Master Sheet:").pack()
    tk.Entry(root, textvariable=master_var, width=60).pack()
    tk.Button(root, text="Browse Master", command=lambda: master_var.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))).pack(pady=5)

    tk.Label(root, text="Working Directory:").pack()
    tk.Entry(root, textvariable=working_dir_var, width=60).pack()
    tk.Button(root, text="Browse Working Dir", command=lambda: working_dir_var.set(filedialog.askdirectory())).pack(pady=5)

    tk.Label(root, text="Group Folder:").pack()
    tk.Entry(root, textvariable=group_var, width=60).pack()
    tk.Button(root, text="Browse Group Folder", command=lambda: group_var.set(filedialog.askdirectory())).pack(pady=5)

    tk.Button(root, text="Run Grade Import", command=run_update, padx=20, pady=10).pack(pady=10)

    root.mainloop()
