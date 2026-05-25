import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from clients.oneok import process_oneok_workbook
from dras_export import generate_dras_files


CLIENTS = {
    "ONEOK": process_oneok_workbook,
}

PIPE_DIAMETERS = [
    "6.625",
    "8.625",
    "10.75",
    "12.75",
    "14.00",
    "16.00",
    "18.00",
]


class ExcelCleanerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Excel Cleaner")
        self.root.geometry("680x280")
        self.root.resizable(False, False)

        self.client_var = tk.StringVar(value="ONEOK")
        self.pipe_diameter_var = tk.StringVar(value=PIPE_DIAMETERS[0])
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

        ttk.Label(frame, text="Pipe Diameter (in)").grid(row=1, column=0, sticky="w", pady=(0, 10))
        self.pipe_diameter_combo = ttk.Combobox(
            frame,
            textvariable=self.pipe_diameter_var,
            values=PIPE_DIAMETERS,
            state="readonly",
            width=40,
        )
        self.pipe_diameter_combo.grid(row=1, column=1, sticky="ew", pady=(0, 10))

        ttk.Label(frame, text="Excel File").grid(row=2, column=0, sticky="w", pady=(0, 10))
        ttk.Entry(
            frame,
            textvariable=self.file_path_var,
            width=60,
        ).grid(row=2, column=1, sticky="ew", pady=(0, 10))

        ttk.Button(frame, text="Browse...", command=self.browse_file).grid(
            row=2, column=2, padx=(8, 0), pady=(0, 10)
        )

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=3, column=1, sticky="w", pady=(10, 0))

        ttk.Button(
            button_frame,
            text="Process File",
            command=self.process_file,
        ).pack(side="left")

        ttk.Button(
            button_frame,
            text="Generate DRAS Files",
            command=self.generate_dras_files,
        ).pack(side="left", padx=(10, 0)
        )

        ttk.Label(
            frame,
            text="Output will be saved as a new processed file.",
        ).grid(row=4, column=1, sticky="w", pady=(12, 0))

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
        pipe_diameter = self.pipe_diameter_var.get().strip()

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

        if pipe_diameter not in PIPE_DIAMETERS:
            messagebox.showerror("Invalid pipe diameter", "Please select a valid pipe diameter.")
            return

        try:
            ext = os.path.splitext(file_path)[1].lower()
            keep_vba = ext in {".xlsm", ".xltm"}

            wb = load_workbook(file_path, keep_vba=keep_vba)

            # Pass pipe diameter into the client processor
            processor(wb, pipe_diameter=float(pipe_diameter))

            output_path = self._build_output_path(file_path, client_name, ext)
            wb.save(output_path)

            messagebox.showinfo(
                "Success",
                f"Processed file saved to:\n{output_path}",
            )


        except PermissionError:
            messagebox.showerror(
                "File in Use",
                "The Excel file is currently open.\n\n"
                "Please close it and try again."
            )    

        except Exception as exc:
            messagebox.showerror(
                "Processing error",
                f"An unexpected error occurred:\n\n{str(exc)}"
            )

    def generate_dras_files(self) -> None:
        file_path = self.file_path_var.get().strip()
        pipe_diameter = self.pipe_diameter_var.get().strip()

        if not file_path:
            messagebox.showwarning("Missing file", "Please select an Excel file first.")
            return

        try:
            output_folder = os.path.join(os.path.dirname(file_path), "DRAS_Output")

            generate_dras_files(
                excel_path=file_path,
                output_folder=output_folder,
                pipe_diameter=float(pipe_diameter),
            )

            messagebox.showinfo(
                "Success",
                f"DRAS files generated in:\n{output_folder}",
            )

        except PermissionError:
            messagebox.showerror(
                "File in Use",
                "The Excel file appears to be open.\n\nPlease close it and try again.",
            )

        except Exception as exc:
            messagebox.showerror(
                "DRAS Export Error",
                f"An error occurred while generating DRAS files:\n\n{str(exc)}",
            )

    @staticmethod
    def _build_output_path(file_path: str, client_name: str, ext: str) -> str:
        base, _ = os.path.splitext(file_path)
        output_ext = ext if ext in {".xlsx", ".xlsm", ".xltm"} else ".xlsx"
        return f"{base}_{client_name.lower()}_processed{output_ext}"


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelCleanerApp(root)
    root.mainloop()
