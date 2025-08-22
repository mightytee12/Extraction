import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import sys
import subprocess

def extract_non_compliant():
    try:
        # Select first file (all compliant)
        file1 = filedialog.askopenfilename(
            title="Select Excel File with ALL Compliant Businesses",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file1:
            return
        
        # Select second file (mixed data)
        file2 = filedialog.askopenfilename(
            title="Select Excel File with Mixed Businesses",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file2:
            return

        # Show progress bar
        progress["value"] = 20
        root.update_idletasks()

        # Load files
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        progress["value"] = 50
        root.update_idletasks()

        # Ensure columns exist
        required_cols = ["BIN", "BUSINESS NAME", "BUSINESS LOCATION", "BARANGAY"]
        for col in required_cols:
            if col not in df2.columns:
                messagebox.showerror("Error", f"Missing column in Data 2: {col}")
                return
        if "BIN" not in df1.columns:
            messagebox.showerror("Error", "Missing column 'BIN' in Data 1")
            return

        # Extract compliant BINs
        compliant_bins = set(df1["BIN"].astype(str))

        # Filter non-compliant records (BIN not in df1)
        non_compliant = df2[~df2["BIN"].astype(str).isin(compliant_bins)].copy()

        # Select required columns
        output = non_compliant[["BIN", "BUSINESS NAME", "BUSINESS LOCATION", "BARANGAY"]]

        progress["value"] = 80
        root.update_idletasks()

        # Save to new Excel with each Barangay in its own sheet
        output_file = os.path.join(os.path.dirname(file2), "SAGOT SA PAGHIHIRAP NI MAICA AT JULIUS.xlsx")
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            if not output.empty:
                for BARANGAY, group in output.groupby("BARANGAY"):
                    if not group.empty:
                        group.to_excel(writer, sheet_name=str(BARANGAY)[:31], index=False)
            else:
                # Fallback if no data
                pd.DataFrame([{"Message": "No Non-Compliant Businesses Found"}]) \
                    .to_excel(writer, sheet_name="Summary", index=False)

        progress["value"] = 100
        root.update_idletasks()

        # Auto open the file
        if sys.platform == "win32":  # Windows
            os.startfile(output_file)
        elif sys.platform == "darwin":  # macOS
            subprocess.call(["open", output_file])
        else:  # Linux
            subprocess.call(["xdg-open", output_file])

        messagebox.showinfo(
            "Success",
            f"Non-compliant businesses extracted!\n\nSaved as:\n{output_file}"
        )

    except Exception as e:
        messagebox.showerror("NAKU PO!!!! MALI!!!!!!!!", str(e))


# GUI
root = tk.Tk()
root.title("PERMITTING MALUPIT")
root.geometry("450x220")

label = tk.Label(root, text="GAWA NI NONOY", font=("Arial", 14))
label.pack(pady=10)

button = tk.Button(root, text="Start Extraction", command=extract_non_compliant,
                   font=("Arial", 12), bg="lightblue")
button.pack(pady=10)

progress = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress.pack(pady=20)

root.mainloop()
