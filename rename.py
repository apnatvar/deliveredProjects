import os
import datetime as dt
from pathlib import Path
from tkinter import Tk, Button, Label, filedialog, messagebox
from openpyxl import load_workbook

replaceDict = {
    "Unspent Amount Payable to Development Work": "Unutilized Amount of Development Work",
    "GIS Claim Payable": "GIS Claim Payable to Ex-Employee",
    "Group Insurance Scheme": "Group Insurance Scheme Payable",
    "GSLI Payable": "GSLI Payable to Ex-Employee",
    "RCM Payable": "GST RCM Payable",
    "TDS on GST Payable": "GST TDS Payable",
    "Tax Collection at Source": "Income Tax(TCS) Payable",
    "Tax Deducted at Source": "Income Tax(TDS)Payable",
    "Other Recovery (Material Loss)": "Recovery agt Material Loss (Staff)",
    "House Building Advance3": "Recovery of House Building Advance",
    "Vehicles Advance": "Recovery of Vehicles Advance",
    "Other Recovery": "Recovery - Other",
    "Security Against Contractors": "Security Deposit from Contractors",
    "Employee's Security": "Security from Employee",
    "Staff Gratuity": "Provision for Staff Gratuity",
    "Plant & Machinery (Others) @ 15%": "Plant & Machinery @ 15%",
    "House Building Advance7": "Advance For House Building to Staff",
    "Vehicle Advance": "Advance For Vehicle to Staff",
    "RCM Input available": "GST RCM Input Available",
    "Other Security": "Security Deposited - Other",
}


def timestamp() -> str:
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def process_xlsx_xlsm(path: Path, rep_dict: dict) -> dict:
    """Process .xlsx or .xlsm ."""
    keep_vba = path.suffix.lower() == ".xlsm"
    wb = load_workbook(filename=str(path), keep_vba=keep_vba, data_only=False)
    counter = {k: 0 for k in rep_dict.keys()}

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if val in rep_dict.keys():
                    cell.value = rep_dict[val]
                    counter[val] += 1

    wb.save(str(path))
    return counter


def process_workbook(file_path: str, rep_dict: dict) -> dict:
    """Detect extension and route to appropriate processor. Returns replaceCountDict."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    ext = path.suffix.lower()
    if ext in (".xlsx", ".xlsm"):
        return process_xlsx_xlsm(path, rep_dict)
    else:
        raise ValueError("Unsupported file type. Use .xlsx, .xlsm")


# ------------------ Tkinter GUI ------------------


class App:  # Tested on Windows Only
    def __init__(self, root):
        root.title("Excel Replace Automation")
        root.geometry("700x240")
        root.resizable(False, False)

        self.label = Label(root, text="Select an Excel file (.xlsx, .xlsm) to process:")
        self.label.pack(pady=10)

        self.button = Button(root, text="Choose File and Run", command=self.run)
        self.button.pack(pady=5)

        self.status = Label(root, text="", fg="gray")
        self.status.pack(pady=5)

    def run(self):
        filetypes = [
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ]
        filename = filedialog.askopenfilename(
            title="Select Excel file", filetypes=filetypes
        )
        if not filename:
            return

        self.status.config(text="Running... please wait.")
        self.button.config(state="disabled")
        root.update_idletasks()

        try:
            counts = process_workbook(filename, replaceDict)
            total_replacements = sum(counts.values())
            nonzero = {k: v for k, v in counts.items() if v}
            lines = [
                f"Replacements complete.",
                f"File: {Path(filename).name}",
                f"Total replacements: {total_replacements}",
            ]
            if nonzero:
                lines.append("Per-key counts:")
                width = max(len(k) for k in nonzero.keys())
                for k, v in sorted(nonzero.items()):
                    lines.append(f"  {k.ljust(width)} : {v}")
            else:
                lines.append("No keys were found.")

            self.status.config(text="Done.")
            messagebox.showinfo("Automation Complete", "\n".join(lines))

        except Exception as e:
            self.status.config(text="Error.")
            messagebox.showerror("Error", f"{type(e).__name__}: {e}")

        finally:
            self.button.config(state="normal")


if __name__ == "__main__":
    root = Tk()
    App(root)
    root.mainloop()
