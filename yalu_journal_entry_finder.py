import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook

APP_TITLE = "YALU Journal Entry Finder"

class JournalFinder:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("420x300")
        self.root.minsize(380, 260)
        self.root.configure(bg="#171b22")
        self.rows = []
        self.headers = []
        self.build_ui()

    def build_ui(self):
        tk.Label(self.root, text="YALU Journal Entry Finder", bg="#171b22",
                 fg="white", font=("Segoe UI", 15, "bold")).pack(pady=(14, 2))
        tk.Label(self.root, text="Prepared By Lakshan", bg="#171b22",
                 fg="#9aa7b7", font=("Segoe UI", 8)).pack(pady=(0, 10))

        tk.Button(self.root, text="Open Excel File", command=self.open_file,
                   bg="#303b4b", fg="white", relief="flat",
                   padx=12, pady=6).pack()

        self.status = tk.Label(self.root, text="No Excel file selected",
                               bg="#171b22", fg="#c2cad4", font=("Segoe UI", 9))
        self.status.pack(pady=8)

        searchrow = tk.Frame(self.root, bg="#171b22")
        searchrow.pack(fill="x", padx=18, pady=4)
        tk.Label(searchrow, text="Entry No.", bg="#171b22", fg="white").pack(side="left")
        self.entry = tk.Entry(searchrow, bg="#252c36", fg="white",
                              insertbackground="white", relief="flat")
        self.entry.pack(side="left", fill="x", expand=True, padx=8, ipady=6)
        tk.Button(searchrow, text="Search", command=self.search,
                  bg="#287a58", fg="white", relief="flat", padx=12).pack(side="right")

        tk.Label(self.root, text="Description", bg="#171b22", fg="#aeb9c7",
                 anchor="w").pack(fill="x", padx=20, pady=(10, 2))
        self.description = tk.Entry(self.root, bg="#252c36", fg="white",
                                    insertbackground="white", relief="flat")
        self.description.pack(fill="x", padx=20, ipady=7)

        tk.Label(self.root, text="Cost Center", bg="#171b22", fg="#aeb9c7",
                 anchor="w").pack(fill="x", padx=20, pady=(8, 2))
        self.cost_center = tk.Entry(self.root, bg="#252c36", fg="white",
                                    insertbackground="white", relief="flat")
        self.cost_center.pack(fill="x", padx=20, ipady=7)

        bottom = tk.Frame(self.root, bg="#171b22")
        bottom.pack(fill="x", padx=20, pady=12)
        tk.Button(bottom, text="Copy Description", command=self.copy_description,
                  bg="#3b4656", fg="white", relief="flat", padx=10, pady=5).pack(side="left")
        tk.Button(bottom, text="Copy Cost Center", command=self.copy_cost,
                  bg="#3b4656", fg="white", relief="flat", padx=10, pady=5).pack(side="left", padx=8)

    def open_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xlsm")])
        if not path:
            return
        try:
            wb = load_workbook(path, data_only=True, read_only=True)
            ws = wb.active
            values = list(ws.iter_rows(values_only=True))
            if not values:
                raise ValueError("The selected worksheet is empty.")
            self.headers = [str(v).strip().lower() if v is not None else "" for v in values[0]]
            self.rows = values[1:]
            self.status.config(text=f"Loaded: {Path(path).name} ({len(self.rows)} rows)")
        except Exception as e:
            messagebox.showerror("Excel error", str(e))

    def find_col(self, names):
        for i, h in enumerate(self.headers):
            if any(n in h for n in names):
                return i
        return None

    def search(self):
        if not self.rows:
            messagebox.showwarning("No file", "Please open an Excel file first.")
            return
        key = self.entry.get().strip().lower()
        if not key:
            return
        entry_col = self.find_col(["entry no", "entry number", "journal no", "journal number", "document no", "doc no", "voucher no"])
        desc_col = self.find_col(["description", "memo", "details", "narration"])
        cost_col = self.find_col(["cost center", "costcentre", "cost centre"])
        if entry_col is None:
            messagebox.showerror("Column not found", "Could not identify the Entry No. column. Check your Excel header row.")
            return
        for row in self.rows:
            value = row[entry_col] if entry_col < len(row) else None
            if value is not None and str(value).strip().lower() == key:
                desc = row[desc_col] if desc_col is not None and desc_col < len(row) else ""
                cost = row[cost_col] if cost_col is not None and cost_col < len(row) else ""
                self.set_entry(self.description, desc)
                self.set_entry(self.cost_center, cost)
                return
        messagebox.showinfo("Not found", f"Entry No. {self.entry.get()} was not found.")

    @staticmethod
    def set_entry(widget, value):
        widget.delete(0, tk.END)
        widget.insert(0, "" if value is None else str(value))

    def copy_description(self):
        self.copy_value(self.description.get())

    def copy_cost(self):
        self.copy_value(self.cost_center.get())

    def copy_value(self, value):
        if not value:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(value)
        self.root.update()
        messagebox.showinfo("Copied", "Copied to clipboard.")

if __name__ == "__main__":
    window = tk.Tk()
    JournalFinder(window)
    window.mainloop()
