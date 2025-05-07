import tkinter as tk
from tkinter import filedialog, messagebox
import csv
from openpyxl import Workbook
import os

def convert_csv_to_excel():
    csv_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not csv_path:
        return  # User cancelled

    try:
        # Read CSV manually without assuming consistent columns
        with open(csv_path, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            data = list(reader)

        # Write to Excel
        excel_path = os.path.splitext(csv_path)[0] + ".xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "CSV Data"

        for row in data:
            ws.append(row)

        wb.save(excel_path)
        messagebox.showinfo("Success", f"Excel file saved:\n{excel_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert:\n{e}")

# Create GUI window
root = tk.Tk()
root.title("CSV to Excel Converter")
root.geometry("300x150")

# Add button
btn = tk.Button(root, text="Convert CSV to Excel", command=convert_csv_to_excel)
btn.pack(expand=True)

# Run the GUI event loop
root.mainloop()

