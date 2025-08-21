import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time

class ComplianceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window settings
        self.title("📑 Business Compliance Checker")
        self.geometry("600x450")
        ctk.set_appearance_mode("light")   # "dark" or "light"
        ctk.set_default_color_theme("blue")

        # File paths
        self.file1 = None
        self.file2 = None

        # Title label
        self.title_label = ctk.CTkLabel(self, text="Business Compliance Dashboard",
                                        font=("Arial", 18, "bold"))
        self.title_label.pack(pady=20)

        # Buttons
        self.btn_file1 = ctk.CTkButton(self, text="Select Compliant List (File 1)", command=self.load_file1)
        self.btn_file1.pack(pady=10)

        self.lbl_file1 = ctk.CTkLabel(self, text="No file selected", text_color="gray")
        self.lbl_file1.pack()

        self.btn_file2 = ctk.CTkButton(self, text="Select Mixed List (File 2)", command=self.load_file2)
        self.btn_file2.pack(pady=10)

        self.lbl_file2 = ctk.CTkLabel(self, text="No file selected", text_color="gray")
        self.lbl_file2.pack()

        # Run button
        self.run_button = ctk.CTkButton(self, text="🔍 Extract Non-Compliant", fg_color="green",
                                        hover_color="darkgreen", command=self.start_processing)
        self.run_button.pack(pady=30)

        # Progress bar
        self.progress = ctk.CTkProgressBar(self, width=400)
        self.progress.set(0)
        self.progress.pack(pady=15)

        self.progress_label = ctk.CTkLabel(self, text="")
        self.progress_label.pack()

    def load_file1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file1 = path
            self.lbl_file1.configure(text=path, text_color="blue")

    def load_file2(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.file2 = path
            self.lbl_file2.configure(text=path, text_color="blue")

    def start_processing(self):
        """Run process in a separate thread so UI doesn’t freeze."""
        thread = threading.Thread(target=self.process)
        thread.start()

    def process(self):
        try:
            if not self.file1 or not self.file2:
                messagebox.showwarning("Missing File", "Please select both files before running.")
                return

            # Start progress
            self.progress_label.configure(text="⏳ Processing...")
            self.progress.set(0)

            # Simulate loading
            for i in range(1, 4):  # Fake steps (reading, filtering, saving)
                time.sleep(0.7)
                self.progress.set(i / 3)
                self.update_idletasks()

            # Actual processing
            df1 = pd.read_excel(self.file1)
            df2 = pd.read_excel(self.file2)

            if "Business Name" not in df1.columns or "Business Name" not in df2.columns or "Status" not in df2.columns:
                messagebox.showerror("Error", "Files must contain 'Business Name' and 'Status' columns.")
                self.progress_label.configure(text="")
                self.progress.set(0)
                return

            non_compliant = df2[df2["Status"].str.lower() == "non-compliant"]
            result = non_compliant[~non_compliant["Business Name"].isin(df1["Business Name"])]

            # Save file
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                       filetypes=[("Excel files", "*.xlsx")])
            if output_file:
                result.to_excel(output_file, index=False)
                self.progress.set(1)
                self.progress_label.configure(text="✅ Done! File saved successfully.")
                messagebox.showinfo("Success", f"Non-compliant businesses saved to:\n{output_file}")
            else:
                self.progress_label.configure(text="❌ Cancelled")
                self.progress.set(0)

        except Exception as e:
            messagebox.showerror("Unexpected Error", str(e))
            self.progress_label.configure(text="❌ Failed")
            self.progress.set(0)


if __name__ == "__main__":
    app = ComplianceApp()
    app.mainloop()
