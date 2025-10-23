"""Aplicación de escritorio para pivotar columnas de un archivo Excel.

Requisitos de paquetes:
- pandas
- openpyxl (motor de pandas para manejar archivos .xlsx)
- tkinter (incluido en la biblioteca estándar de Python en la mayoría de las distribuciones)
- os (para gestionar rutas de archivo)
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Tuple

import pandas as pd

# ---------------------------------------------------------------------------
# Configuración de la pantalla inicial (splash screen).
# Cambie la ruta en `SPLASH_IMAGE_PATH` para utilizar la imagen que prefiera.
# El programa intentará cargar la imagen; si no la encuentra o falla, mostrará
# un mensaje de texto durante los 5 segundos definidos en `SPLASH_DURATION_MS`.
# ---------------------------------------------------------------------------
SPLASH_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "splash.png")
SPLASH_DURATION_MS = 5000


class ScrollableCheckboxFrame(ttk.Frame):
    """Frame con scroll vertical para albergar widgets de selección."""

    def __init__(self, master: tk.Widget) -> None:
        super().__init__(master)
        canvas = tk.Canvas(self, borderwidth=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner.bind(
            "<Configure>", lambda event: canvas.configure(scrollregion=canvas.bbox("all"))
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
        self.geometry("780x620")

        # DataFrame cargado desde el archivo Excel seleccionado.
        self.dataframe: pd.DataFrame | None = None
        self.file_path: str | None = None

        # Diccionarios que relacionan el nombre de la columna con el estado del checkbox.
        self.fixed_vars: Dict[str, tk.BooleanVar] = {}
        self.pivot_vars: Dict[str, tk.BooleanVar] = {}

        # Variables y menús desplegables para los emparejamientos entre columnas pivotables.
        self.pair_vars: Dict[str, tk.StringVar] = {}
        self.pair_menus: Dict[str, tk.OptionMenu] = {}

        self._build_widgets()

    def _build_widgets(self) -> None:
        """Crea la interfaz gráfica y configura los elementos."""

        instructions = (
            "1. Seleccione un archivo Excel (.xlsx).\n"
            "2. Marque las columnas que quedarán fijas y las que se pivotarán.\n"
            "   Puede emparejar columnas pivotables con el menú desplegable correspondiente.\n"
            "3. Pulse \"Transformar y Guardar\" para generar el archivo resultado.xlsx."
        )
        ttk.Label(self, text=instructions, justify="left").pack(
            anchor="w", padx=10, pady=(10, 5)
        )

        select_button = ttk.Button(
            self, text="Cargar archivo Excel", command=self.load_excel
        )
        select_button.pack(padx=10, pady=5, anchor="w")

        self.file_label = ttk.Label(self, text="Ningún archivo seleccionado")
        self.file_label.pack(anchor="w", padx=10, pady=(0, 10))

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        fixed_frame = ttk.Labelframe(container, text="Columnas fijas")
        fixed_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        pivot_frame = ttk.Labelframe(container, text="Columnas pivotables y emparejamientos")
        pivot_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        self.fixed_checkbox_frame = ScrollableCheckboxFrame(fixed_frame)
        self.fixed_checkbox_frame.pack(fill="both", expand=True)

        self.pivot_checkbox_frame = ScrollableCheckboxFrame(pivot_frame)
        self.pivot_checkbox_frame.pack(fill="both", expand=True)

        generate_button = ttk.Button(
            self,
            text="Transformar y Guardar",
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

    def _populate_checkboxes(self, columns: List[str]) -> None:
        """Crea un checkbox por columna para los apartados de fijas y pivotables."""

        # Limpiar cualquier selección previa eliminando los widgets.
        for child in self.fixed_checkbox_frame.inner.winfo_children():
            child.destroy()
        for child in self.pivot_checkbox_frame.inner.winfo_children():
            child.destroy()

        self.fixed_vars.clear()
        self.pivot_vars.clear()
        self.pair_vars.clear()
        self.pair_menus.clear()

        for column in columns:
            fixed_var = tk.BooleanVar(value=False)
            self.fixed_vars[column] = fixed_var
            ttk.Checkbutton(
                self.fixed_checkbox_frame.inner, text=column, variable=fixed_var
            ).pack(anchor="w", padx=5, pady=2)

            pivot_var = tk.BooleanVar(value=False)
            self.pivot_vars[column] = pivot_var

            row_frame = ttk.Frame(self.pivot_checkbox_frame.inner)
            row_frame.pack(fill="x", padx=5, pady=2)

            ttk.Checkbutton(
                row_frame,
                text=column,
                variable=pivot_var,
                command=lambda c=column: self._on_pivot_toggle(c),
            ).pack(side="left", anchor="w")

            pair_var = tk.StringVar(value="Ninguno")
            self.pair_vars[column] = pair_var

            option_menu = tk.OptionMenu(
                row_frame,
                pair_var,
                "Ninguno",
            )
            option_menu.configure(state="disabled")
            option_menu.pack(side="right", padx=(10, 0))
            self.pair_menus[column] = option_menu

        # Ajustar las opciones iniciales de los menús desplegables.
        self._refresh_pair_options()

    def _selected_pivot_columns(self) -> List[str]:
        """Devuelve las columnas marcadas como pivotables conservando el orden original."""

        return [column for column, var in self.pivot_vars.items() if var.get()]

    def _on_pivot_toggle(self, column: str) -> None:
        """Habilita o deshabilita el menú de emparejamiento según el estado del checkbox."""

        is_selected = self.pivot_vars[column].get()
        menu = self.pair_menus[column]
        if is_selected:
            menu.configure(state="normal")
        else:
            # Si la columna deja de ser pivotable se limpia su emparejamiento y el de terceros.
            for other, var in self.pair_vars.items():
                if var.get() == column:
                    self._set_pair(other, "Ninguno")
            self._set_pair(column, "Ninguno")
            menu.configure(state="disabled")
        self._refresh_pair_options()

    def _set_pair(self, column: str, value: str, update_relations: bool = True) -> None:
        """Actualiza el valor del menú de emparejamiento, gestionando relaciones si corresponde."""

        previous = self.pair_vars[column].get()
        if previous == value:
            return
        self.pair_vars[column].set(value)
        if update_relations:
            self._apply_pair_change(column, previous, value)

    def _apply_pair_change(self, column: str, previous: str, new_partner: str) -> None:
        """Sincroniza los emparejamientos evitando referencias inconsistentes."""

        # Romper la relación anterior si existía.
        if previous != "Ninguno" and previous in self.pair_vars:
            if self.pair_vars[previous].get() == column:
                self._set_pair(previous, "Ninguno", update_relations=False)

        if new_partner == "Ninguno":
            return

        # Validar que la nueva pareja siga marcada como pivotable.
        if new_partner not in self.pair_vars or not self.pivot_vars[new_partner].get():
            messagebox.showwarning(
                "Emparejamiento inválido",
                "La columna seleccionada como pareja no está marcada como pivotable.",
            )
            self._set_pair(column, "Ninguno", update_relations=False)
            return

        partner_current = self.pair_vars[new_partner].get()
        if partner_current not in ("Ninguno", column):
            other = partner_current
            self._set_pair(new_partner, "Ninguno", update_relations=False)
            if other in self.pair_vars and self.pair_vars[other].get() == new_partner:
                self._set_pair(other, "Ninguno", update_relations=False)

        if self.pair_vars[new_partner].get() != column:
            self._set_pair(new_partner, column, update_relations=False)

    def _on_option_menu_select(self, column: str, selection: str) -> None:
        """Gestiona la selección de pareja realizada por el usuario."""

        self._set_pair(column, selection)
        self._refresh_pair_options()

    def _refresh_pair_options(self) -> None:
        """Actualiza dinámicamente los valores disponibles en cada menú de emparejamiento."""

        selected_columns = self._selected_pivot_columns()
        for column, menu in self.pair_menus.items():
            menu_widget = menu["menu"]
            menu_widget.delete(0, "end")

            options = ["Ninguno"] + [c for c in selected_columns if c != column]
            for option in options:
                menu_widget.add_command(
                    label=option,
                    command=lambda opt=option, col=column: self._on_option_menu_select(col, opt),
                )

            current = self.pair_vars[column].get()
            if current not in options:
                self._set_pair(column, "Ninguno")

            if self.pivot_vars[column].get():
                menu.configure(state="normal")
            else:
                menu.configure(state="disabled")

    def _build_groups(self, pivot_columns: List[str]) -> List[Tuple[str, ...]]:
        """Construye las tuplas de columnas que se agruparán en la transformación."""

        groups: List[Tuple[str, ...]] = []
        visited = set()
        for column in pivot_columns:
            if column in visited:
                continue
            partner = self.pair_vars[column].get()
            if (
                partner not in (None, "Ninguno")
                and partner in pivot_columns
                and (partner_var := self.pair_vars.get(partner)) is not None
                and partner_var.get() == column
            ):
                groups.append((column, partner))
                visited.update({column, partner})
            else:
                groups.append((column,))
                visited.add(column)
        return groups

    def generate_pivoted_file(self) -> None:
        """Genera un nuevo archivo Excel con la estructura pivotada y lo guarda."""

        if self.dataframe is None:
            messagebox.showwarning(
                "Sin datos",
                "Debe seleccionar un archivo Excel antes de generar el resultado.",
            )
            return

        fixed_columns = [column for column, var in self.fixed_vars.items() if var.get()]
        pivot_columns = self._selected_pivot_columns()

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

        groups = self._build_groups(pivot_columns)
        if not groups:
            messagebox.showwarning(
                "Sin columnas",
                "No se detectaron columnas para transformar.",
            )
            return

        transformed_rows = []
        for _, row in self.dataframe.iterrows():
            fixed_values = row[fixed_columns].to_dict()
            for group in groups:
                new_row = {column: fixed_values[column] for column in fixed_columns}
                base_column = group[0]
                new_row["Columna_base"] = base_column
                new_row["Valor_A"] = row[base_column]
                if len(group) == 2:
                    partner_column = group[1]
                    new_row["Valor_B"] = row[partner_column]
                else:
                    new_row["Valor_B"] = pd.NA
                transformed_rows.append(new_row)

        if not transformed_rows:
            messagebox.showinfo(
                "Sin resultados",
                "No se generaron filas en el resultado. Verifique los datos de origen.",
            )
            return

        result_df = pd.DataFrame(transformed_rows)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, "resultado.xlsx")

        try:
            result_df.to_excel(output_path, index=False)
        except Exception as exc:
            messagebox.showerror(
                "Error al guardar",
                f"No fue posible guardar el archivo de resultado. Detalle: {exc}",
            )
            return

        messagebox.showinfo(
            "Proceso completado",
            f"El archivo se generó correctamente en: {output_path}",
        )

    def mainloop(self, n: int = 0) -> None:  # type: ignore[override]
        """Sobrescritura para asegurar refresco de menús ante cualquier cambio."""

        self._refresh_pair_options()
        super().mainloop(n)


def show_splash_screen() -> None:
    """Muestra una pantalla inicial con imagen configurable durante 5 segundos."""

    splash_root = tk.Tk()
    splash_root.overrideredirect(True)
    splash_root.configure(bg="white")
    splash_root.resizable(False, False)

    splash_image = None
    if SPLASH_IMAGE_PATH and os.path.isfile(SPLASH_IMAGE_PATH):
        try:
            splash_image = tk.PhotoImage(file=SPLASH_IMAGE_PATH)
        except tk.TclError as exc:
            print(
                "Advertencia: no se pudo cargar la imagen del splash. "
                f"Detalle: {exc}"
            )

    if splash_image is not None:
        width = splash_image.width()
        height = splash_image.height()
        label = tk.Label(splash_root, image=splash_image, borderwidth=0)
        label.image = splash_image  # Se guarda la referencia para evitar que se libere.
    else:
        width = 400
        height = 300
        label = tk.Label(
            splash_root,
            text=(
                "Configure `SPLASH_IMAGE_PATH` con la ruta de su imagen\n"
                "para personalizar esta pantalla inicial."
            ),
            font=("Segoe UI", 12),
            bg="white",
            justify="center",
            padx=20,
            pady=20,
        )

    label.pack(fill="both", expand=True)

    splash_root.update_idletasks()
    screen_width = splash_root.winfo_screenwidth()
    screen_height = splash_root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    splash_root.geometry(f"{width}x{height}+{x}+{y}")

    splash_root.after(SPLASH_DURATION_MS, splash_root.destroy)
    splash_root.mainloop()


def main() -> None:
    """Punto de entrada principal de la aplicación."""

    show_splash_screen()
    app = PivotApp()
    app.mainloop()


if __name__ == "__main__":
    main()
