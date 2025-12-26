# -*- coding: utf-8 -*-
"""
Aplicación de escritorio para ejecutar el exportador de Presto.
"""

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from presto_commercial import export_to_excel


class PrestoApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Presto Commercial")
        self.geometry("640x360")
        self.resizable(False, False)

        self.excel_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.output_path_var = tk.StringVar(value="resource_totals.xlsx")
        self.status_var = tk.StringVar(value="Listo.")

        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 12, "pady": 8}

        frame = ttk.Frame(self)
        frame.pack(fill=tk.BOTH, expand=True, **padding)

        title = ttk.Label(frame, text="Exportar recursos desde Presto", font=("Segoe UI", 14, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))

        ttk.Label(frame, text="Archivo Excel:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.excel_path_var, width=56).grid(row=1, column=1, sticky="w")
        ttk.Button(frame, text="Buscar", command=self._select_excel).grid(row=1, column=2, sticky="w")

        ttk.Label(frame, text="Hoja (opcional):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.sheet_name_var, width=56).grid(row=2, column=1, sticky="w")

        ttk.Label(frame, text="Salida:").grid(row=3, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_path_var, width=56).grid(row=3, column=1, sticky="w")
        ttk.Button(frame, text="Guardar como", command=self._select_output).grid(row=3, column=2, sticky="w")

        self.run_button = ttk.Button(frame, text="Procesar", command=self._run_export)
        self.run_button.grid(row=4, column=1, sticky="e", pady=(16, 0))

        ttk.Separator(frame).grid(row=5, column=0, columnspan=3, sticky="ew", pady=12)
        ttk.Label(frame, textvariable=self.status_var, foreground="#1a1a1a").grid(
            row=6, column=0, columnspan=3, sticky="w"
        )

        frame.columnconfigure(1, weight=1)

    def _select_excel(self) -> None:
        filename = filedialog.askopenfilename(
            title="Selecciona el Excel de Presto",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls")],
        )
        if filename:
            self.excel_path_var.set(filename)
            if not self.output_path_var.get():
                self.output_path_var.set(os.path.join(os.path.dirname(filename), "resource_totals.xlsx"))

    def _select_output(self) -> None:
        filename = filedialog.asksaveasfilename(
            title="Guardar como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if filename:
            self.output_path_var.set(filename)

    def _set_busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.run_button.configure(state=state)

    def _run_export(self) -> None:
        excel_path = self.excel_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        sheet_name = self.sheet_name_var.get().strip() or None

        if not excel_path:
            messagebox.showwarning("Datos incompletos", "Debe seleccionar un archivo Excel.")
            return
        if not output_path:
            messagebox.showwarning("Datos incompletos", "Debe indicar la ruta de salida.")
            return

        self.status_var.set("Procesando...")
        self._set_busy(True)

        def _task() -> None:
            try:
                base, ext = os.path.splitext(output_path)
                final_output = output_path if ext.lower() in {".xlsx", ".xlsm", ".xls"} else base + ".xlsx"
                export_to_excel(excel_path, final_output, sheet_name=sheet_name)
            except Exception as exc:
                self.after(0, lambda: self._handle_error(exc))
                return
            self.after(0, lambda: self._handle_success(final_output))

        threading.Thread(target=_task, daemon=True).start()

    def _handle_error(self, exc: Exception) -> None:
        self.status_var.set("Error al procesar el archivo.")
        self._set_busy(False)
        messagebox.showerror("Error", f"No se pudo completar el proceso.\n\n{exc}")

    def _handle_success(self, output_path: str) -> None:
        self.status_var.set("Proceso completado correctamente.")
        self._set_busy(False)
        messagebox.showinfo("Listo", f"Archivo exportado en:\n{output_path}")


def main() -> None:
    app = PrestoApp()
    app.mainloop()


if __name__ == "__main__":
    main()
