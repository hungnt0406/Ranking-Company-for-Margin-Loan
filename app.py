import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


def run_processing():
    import pandas as pd
    import model  # your model.py file
    try:
        # Step 1: Ask user to choose Excel file
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return  # user cancelled

        # Step 2: Load all sheet names
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # Step 3: Ask user to choose a sheet
        sheet_choice = simpledialog.askstring(
            "Select Sheet",
            f"Available sheets:\n{', '.join(sheet_names)}\n\nType the sheet name to use:"
        )
        if not sheet_choice:
            return

        if sheet_choice not in sheet_names:
            messagebox.showerror("Error", f"Sheet '{sheet_choice}' not found in file!")
            return

        # Step 4: Read the chosen sheet
        df = pd.read_excel(file_path, sheet_name=sheet_choice)

        # Step 6: Ask user where to save output
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save Result As"
        )
        if not save_path:
            return

        # Step 7: Save the result
        model.to_excel(df,save_path)

        messagebox.showinfo("Success", f"File processed successfully!\nSaved to:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Something went wrong:\n{e}")

# --- GUI setup ---
root = tk.Tk()
root.title("Excel Auto Scoring Tool")
root.geometry("420x220")

tk.Label(root, text="📊 Excel Auto Scoring App", font=("Arial", 14, "bold")).pack(pady=20)
tk.Button(root, text="Select Excel and Run", command=run_processing, width=22, height=2).pack(pady=10)
tk.Label(root, text="Created by You", font=("Arial", 8)).pack(side="bottom", pady=10)

root.mainloop()
