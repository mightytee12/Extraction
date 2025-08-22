import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox

# Load Excel file
df = pd.read_excel("GHGClearanceDatabase2025.xlsx")  # update with your file name

def search_data():
    query = entry.get().strip().lower()
    results = df[
        df.apply(lambda row: query in str(row['BIN']).lower() 
                 or query in str(row['BUSINESS NAME']).lower(), axis=1)
    ]

    # Clear old results
    for row in tree.get_children():
        tree.delete(row)

    # Insert new results
    if not results.empty:
        for _, row in results.iterrows():
            tree.insert("", tk.END, values=(row['BIN'], row['Business Name'], row['Business Location'], row['Status']))
    else:
        messagebox.showinfo("No Results", "No matching record found.")

# GUI Setup
root = tk.Tk()
root.title("Database Search")
root.geometry("800x400")

# Search box
tk.Label(root, text="Search BIN or BUSINESS NAME:").pack(pady=5)
entry = tk.Entry(root, width=50)
entry.pack(pady=5)
tk.Button(root, text="Search", command=search_data).pack(pady=5)

# Results table
cols = ("BIN", "BUSINESS NAME", "BUSINESS LOCATION", "STATUS")
tree = ttk.Treeview(root, columns=cols, show="headings")
for col in cols:
    tree.heading(col, text=col)
    tree.column(col, width=150)
tree.pack(expand=True, fill="both", pady=10)

root.mainloop()