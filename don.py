import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from openpyxl import Workbook


files = []

def clean_text(text):
    return str(text).strip().lower()


def parse_amount(amount_text):
    amount_text = str(amount_text).replace("$", "").replace(",", "").strip()
    try:
        return float(amount_text)
    except:
        return 0.0


def add_files(file_list):
    for file in file_list:
        file = file.strip("{}")
        if file.lower().endswith(".csv") and file not in files:
            files.append(file)
            listbox.insert(tk.END, file)


def choose_files():
    selected = filedialog.askopenfilenames(filetypes=[("CSV files", "*.csv")])
    add_files(selected)


def drop_files(event):
    dropped = root.tk.splitlist(event.data)
    add_files(dropped)


def clear_files():
    files.clear()
    listbox.delete(0, tk.END)


def merge_files():
    if not files:
        messagebox.showwarning("Warning", "Please add CSV files first.")
        return

    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")]
    )

    if not output_file:
        return

    try:
        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Merged"
        ws1.append(["Transaction Date", "Description", "Category", "Amount"])

        ws2 = wb.create_sheet("Category Totals")
        ws2.append(["Category", "Total Expense"])

        category_totals = {}

        for file in files:
            with open(file, "r", newline="", encoding="utf-8-sig") as f:
                reader = csv.reader(f)

                header = next(reader, None)  # skip header row
                if not header:
                    continue

                for row in reader:
                    if len(row) < 7:
                        continue

                    # Original column E = Type
                    if "payment" in clean_text(row[4]):
                        continue

                    date = row[0].strip()           # Transaction Date
                    category = row[3].strip()       # Category
                    amount = parse_amount(row[5])   # Amount
                    details = row[2].strip()        # Description

                    ws1.append([date, details, category, amount])

                    category_totals[category] = category_totals.get(category, 0) + amount

        for category, total in sorted(category_totals.items()):
            ws2.append([category, total])

        wb.save(output_file)
        messagebox.showinfo("Done", "Excel file created successfully.")

    except Exception as e:
        messagebox.showerror("Error", str(e))


root = TkinterDnD.Tk()
root.title("CSV Merger")
root.geometry("600x400")

tk.Label(root, text="Drag CSV files here or click button").pack(pady=10)

listbox = tk.Listbox(root, width=80, height=15)
listbox.pack(padx=10, pady=10, fill="both", expand=True)

listbox.drop_target_register(DND_FILES)
listbox.dnd_bind("<<Drop>>", drop_files)

frame = tk.Frame(root)
frame.pack(pady=10)

tk.Button(frame, text="Choose Files", command=choose_files).pack(side="left", padx=5)
tk.Button(frame, text="Clear", command=clear_files).pack(side="left", padx=5)
tk.Button(frame, text="Merge", command=merge_files).pack(side="left", padx=5)

root.mainloop()
