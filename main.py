import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from clients.oneok import process_oneok_workbook


CLIENTS = {
    "ONEOK": process_oneok_workbook,
}


class ExcelCleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Cleaner")
        self.root.geometry("650x240")
        self.root.resizable(False, False)

        self.client_var = tk.StringVar(value="ONEOK")
        self.file_path_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Client").grid(row=0, column=0, sticky="w", pady=(0, 10))

        self.client_combo = ttk.Combobox(
            frame,
            textvariable=self.client_var,
            values=list(CLIENTS.keys()),
            state="readonly",
            width=40,
        )
        self.client_combo.grid(row=0, column=1, sticky="ew", pady=(0, 10))

        ttk.Label(frame, text="Excel File").grid(row=1, column=0, sticky="w", pady=(0, 10))

        ttk.Entry(
            frame,
            textvariable=self.file_path_var,
            width=60,
        ).grid(row=1, column=1, sticky="ew", pady=(0, 10))

        ttk.Button(frame, text="Browse...", command=self.browse_file).grid(
            row=1, column=2, padx=(8, 0), pady=(0, 10)
        )

        ttk.Button(frame, text="Process File", command=self.process_file).grid(
            row=2, column=1, sticky="w", pady=(10, 0)
        )

        ttk.Label(
            frame,
            text="Output will be saved as a new processed file.",
        ).grid(row=3, column=1, sticky="w", pady=(12, 0))

        frame.columnconfigure(1, weight=1)

    def browse_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[
                ("Excel Workbook", "*.xlsx"),
                ("Excel Macro-Enabled Workbook", "*.xlsm"),
                ("Excel Macro-Enabled Template", "*.xltm"),
                ("All files", "*.*"),
            ],
        )

        if file_path:
            self.file_path_var.set(file_path)

    def process_file(self) -> None:
        file_path = self.file_path_var.get().strip()
        client_name = self.client_var.get().strip()

        if not file_path:
            messagebox.showwarning("Missing file", "Please select an Excel file.")
            return

        if not os.path.exists(file_path):
            messagebox.showerror("File not found", "The selected file does not exist.")
            return

        processor = CLIENTS.get(client_name)
        if processor is None:
            messagebox.showerror("Invalid client", f"Unknown client: {client_name}")
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            keep_vba = ext in {".xlsm", ".xltm"}

            wb = load_workbook(file_path, keep_vba=keep_vba)
            processor(wb)

            output_path = self._build_output_path(file_path, client_name, ext)
            wb.save(output_path)

            messagebox.showinfo(
                "Success",
                f"Processed file saved to:\n{output_path}",
            )

        except Exception as exc:
            messagebox.showerror("Processing error", str(exc))

    @staticmethod
    def _build_output_path(file_path: str, client_name: str, ext: str) -> str:
        base, _ = os.path.splitext(file_path)
        output_ext = ext if ext in {".xlsx", ".xlsm", ".xltm"} else ".xlsx"
        return f"{base}_{client_name.lower()}_processed{output_ext}"


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()
