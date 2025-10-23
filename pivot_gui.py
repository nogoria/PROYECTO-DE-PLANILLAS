"""Aplicación de escritorio para transformar columnas de un Excel mediante pivotaje.

Requisitos de paquetes:
- pandas
- openpyxl (motor de pandas para escribir/leer archivos .xlsx)
- tkinter (incluido en la biblioteca estándar de Python en la mayoría de las distribuciones)
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict

import pandas as pd


class ScrollableCheckboxFrame(ttk.Frame):
    """Frame con scroll vertical para albergar widgets de selección."""

    def __init__(self, master: tk.Widget) -> None:
        super().__init__(master)
        canvas = tk.Canvas(self, borderwidth=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")


class PivotApp(tk.Tk):
    """Ventana principal de la aplicación."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Pivotador de columnas")
        self.geometry("700x600")

        # DataFrame cargado desde el archivo Excel seleccionado.
        self.dataframe: pd.DataFrame | None = None
        self.file_path: str | None = None

        # Diccionarios que relacionan el nombre de la columna con el estado del checkbox.
        self.fixed_vars: Dict[str, tk.BooleanVar] = {}
        self.pivot_vars: Dict[str, tk.BooleanVar] = {}

        self._build_widgets()

    def _build_widgets(self) -> None:
        """Crea la interfaz gráfica y configura los elementos."""

        instructions = (
            "1. Seleccione un archivo Excel (.xlsx).\n"
            "2. Elija qué columnas permanecerán fijas y cuáles se pivotarán.\n"
            "3. Pulse \"Generar archivo\" para guardar el resultado."
        )
        ttk.Label(self, text=instructions, justify="left").pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        select_button = ttk.Button(
            self, text="Seleccionar archivo Excel", command=self.load_excel
        )
        select_button.pack(padx=10, pady=5, anchor="w")

        self.file_label = ttk.Label(self, text="Ningún archivo seleccionado")
        self.file_label.pack(anchor="w", padx=10, pady=(0, 10))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        fixed_frame = ttk.Labelframe(container, text="Columnas fijas")
        fixed_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        pivot_frame = ttk.Labelframe(container, text="Columnas a pivotar")
        pivot_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.fixed_checkbox_frame = ScrollableCheckboxFrame(fixed_frame)
        self.fixed_checkbox_frame.pack(fill="both", expand=True)

        self.pivot_checkbox_frame = ScrollableCheckboxFrame(pivot_frame)
        self.pivot_checkbox_frame.pack(fill="both", expand=True)

        generate_button = ttk.Button(
            self,
            text="Generar archivo",
            command=self.generate_pivoted_file,
        )
        generate_button.pack(pady=10)

    def load_excel(self) -> None:
        """Solicita al usuario un archivo Excel y carga los encabezados en la interfaz."""

        filetypes = [("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        path = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=filetypes)
        if not path:
            return  # El usuario canceló la selección.

        try:
            dataframe = pd.read_excel(path, sheet_name=0)
        except FileNotFoundError:
            messagebox.showerror(
                "Archivo no encontrado",
                "El archivo seleccionado no existe. Por favor, intente nuevamente.",
            )
            return
        except Exception as exc:  # Captura errores como falta de permisos o formato inválido.
            messagebox.showerror(
                "Error al leer el archivo",
                f"No fue posible leer el Excel seleccionado. Detalle: {exc}",
            )
            return

        if dataframe.empty:
            messagebox.showwarning(
                "Archivo vacío",
                "El archivo cargado no contiene filas. Seleccione un archivo con datos.",
            )
            return

        self.dataframe = dataframe
        self.file_path = path
        self.file_label.config(text=f"Archivo seleccionado: {path}")

        self._populate_checkboxes(list(dataframe.columns))

    def _populate_checkboxes(self, columns: list[str]) -> None:
        """Crea un checkbox por columna para los apartados de fijas y pivotables."""

        # Limpiar cualquier selección previa eliminando los widgets.
        for child in self.fixed_checkbox_frame.inner.winfo_children():
            child.destroy()
        for child in self.pivot_checkbox_frame.inner.winfo_children():
            child.destroy()

        self.fixed_vars.clear()
        self.pivot_vars.clear()

        for column in columns:
            fixed_var = tk.BooleanVar(value=False)
            pivot_var = tk.BooleanVar(value=False)
            self.fixed_vars[column] = fixed_var
            self.pivot_vars[column] = pivot_var

            ttk.Checkbutton(
                self.fixed_checkbox_frame.inner, text=column, variable=fixed_var
            ).pack(anchor="w", padx=5, pady=2)

            ttk.Checkbutton(
                self.pivot_checkbox_frame.inner, text=column, variable=pivot_var
            ).pack(anchor="w", padx=5, pady=2)

    def generate_pivoted_file(self) -> None:
        """Genera un nuevo archivo Excel con la estructura pivotada."""

        if self.dataframe is None:
            messagebox.showwarning(
                "Sin datos",
                "Debe seleccionar un archivo Excel antes de generar el resultado.",
            )
            return

        fixed_columns = [column for column, var in self.fixed_vars.items() if var.get()]
        pivot_columns = [column for column, var in self.pivot_vars.items() if var.get()]

        if not fixed_columns:
            messagebox.showwarning(
                "Selección incompleta",
                "Seleccione al menos una columna fija.",
            )
            return

        if not pivot_columns:
            messagebox.showwarning(
                "Selección incompleta",
                "Seleccione al menos una columna a pivotar.",
            )
            return

        repeated_columns = set(fixed_columns) & set(pivot_columns)
        if repeated_columns:
            messagebox.showwarning(
                "Selección inválida",
                "Una columna no puede ser fija y pivotada a la vez. Revise su selección.",
            )
            return

        # Construcción manual del DataFrame pivotado según la especificación solicitada.
        transformed_rows = []
        for _, row in self.dataframe.iterrows():
            fixed_values = row[fixed_columns].to_dict()
            for pivot_column in pivot_columns:
                new_row = fixed_values.copy()
                new_row["Columna_pivotada"] = pivot_column
                new_row["Valor"] = row[pivot_column]
                transformed_rows.append(new_row)

        if not transformed_rows:
            messagebox.showinfo(
                "Sin resultados",
                "No se generaron filas en el resultado. Verifique los datos de origen.",
            )
            return

        result_df = pd.DataFrame(transformed_rows)

        default_name = "resultado.xlsx"
        initialdir = os.path.dirname(self.file_path) if self.file_path else None

        save_path = filedialog.asksaveasfilename(
            title="Guardar resultado",
            defaultextension=".xlsx",
            filetypes=[("Archivos de Excel", "*.xlsx")],
            initialfile=default_name,
            initialdir=initialdir,
        )

        if not save_path:
            return  # El usuario canceló el guardado.

        try:
            result_df.to_excel(save_path, index=False)
        except Exception as exc:
            messagebox.showerror(
                "Error al guardar",
                f"No fue posible guardar el archivo de resultado. Detalle: {exc}",
            )
            return

        messagebox.showinfo(
            "Proceso completado",
            f"El archivo se generó correctamente en: {save_path}",
        )


def main() -> None:
    """Punto de entrada principal de la aplicación."""

    app = PivotApp()
    app.mainloop()


if __name__ == "__main__":
    main()
