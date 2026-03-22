import csv
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES


files = []


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
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")]
    )

    if not output_file:
        return

    try:
        with open(output_file, "w", newline="", encoding="utf-8-sig") as out:
            writer = csv.writer(out)
            wrote_header = False

            for file in files:
                with open(file, "r", newline="", encoding="utf-8-sig") as f:
                    reader = csv.reader(f)

                    header = next(reader, None)  # skip first row
                    if not header:
                        continue

                    if not wrote_header:
                        writer.writerow(header[:4] + header[5:])  # remove column E
                        wrote_header = True

                    for row in reader:
                        if len(row) < 5:
                            continue

                        if "payment" in row[4].strip().lower():
                            continue

                        writer.writerow(row[:4] + row[5:])  # remove column E

        messagebox.showinfo("Done", "Merged successfully.")

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