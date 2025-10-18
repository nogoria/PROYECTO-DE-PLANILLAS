import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


FIELDS = [
    "ID",
    "Número Caso BE/ Radicado Cia/ Control",
    "No. De Poliza",
    "NIT",
    "Nombre Grupo Económico",
    "Nombre Cliente",
    "Ramo",
    "Aseguradora",
    "Modelo de atención",
    "Director Ops",
    "Director Comercial",
    "Nombre Ejecutivo",
    "Canal recepción",
    "Tipo de Trámite",
    "Sub Tipo de Trámite",
    "Identificacion Asegurado Principal",
    "Número documento solicitante",
    "Nombre Solicitante",
    "Parentesco",
    "Fecha radicación a wtw",
    "Fecha Seguimiento 1",
    "Fecha Seguimiento 2",
    "Fecha Seguimiento 3",
    "Observaciones / Seguimiento",
    "Fecha de entrega información cliente",
    "Fecha Radicación Trámite o GBC",
    "Fecha esperada de respuesta",
    "Fecha Respuesta Aseguradora / GBC",
    "Fecha Respuesta Final al Cliente",
    "Estado del Trámite",
    "Subestado del Trámite",
    "Pendiente en Cabeza de",
    "Observación Final de Cierre",
    "Días hasta inicio del Ejecutivo",
    "Calidad respuesta Ejecutivo",
    "Días radicación a respuesta aseguradora / GBC",
    "Calidad respuesta Aseguradora / GBC",
    "Días totales del trámite",
    "Calidad del Cierre",
]


class DataManager:
    def __init__(self) -> None:
        self.records: list[dict[str, str]] = []
        self.display_fields: list[str] = FIELDS[:5]
        self.external_data: list[dict[str, str]] = []
        self.external_key: str | None = None
        self.external_fields: list[str] = []

    def add_record(self, record: dict[str, str]) -> None:
        record_copy = record.copy()
        if self.external_data and self.external_key:
            key_value = record_copy.get(self.external_key, "")
            match = next(
                (
                    row
                    for row in self.external_data
                    if row.get(self.external_key, "") == key_value
                ),
                None,
            )
            if match:
                for field in self.external_fields:
                    record_copy.setdefault(field, match.get(field, ""))
        self.records.append(record_copy)

    def set_display_fields(self, fields: list[str]) -> None:
        if fields:
            self.display_fields = fields

    def set_external_configuration(
        self, *, key_field: str, fields: list[str], data: list[dict[str, str]]
    ) -> None:
        self.external_key = key_field
        self.external_fields = fields
        self.external_data = data


class ScrollableForm(ttk.Frame):
    def __init__(self, master: tk.Widget, *, row_height: int = 2):
        super().__init__(master)
        canvas = tk.Canvas(self, borderwidth=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)
        self.inner.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self.inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.row_height = row_height


class DataEntryTab(ttk.Frame):
    def __init__(self, master: ttk.Notebook, data_manager: DataManager, on_save: callable):
        super().__init__(master)
        self.data_manager = data_manager
        self.on_save = on_save
        self.entries: dict[str, tk.Entry] = {}

        scroll_form = ScrollableForm(self)
        scroll_form.pack(fill="both", expand=True, padx=10, pady=10)

        inner = scroll_form.inner
        for index, field in enumerate(FIELDS):
            label = ttk.Label(inner, text=field, anchor="w")
            entry = ttk.Entry(inner)
            label.grid(row=index, column=0, sticky="w", padx=5, pady=2)
            entry.grid(row=index, column=1, sticky="ew", padx=5, pady=2)
            inner.grid_columnconfigure(1, weight=1)
            self.entries[field] = entry

        button_frame = ttk.Frame(self)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))

        save_button = ttk.Button(button_frame, text="Guardar Registro", command=self.save_record)
        save_button.pack(side="left")

        clear_button = ttk.Button(button_frame, text="Limpiar", command=self.clear_form)
        clear_button.pack(side="left", padx=5)

    def clear_form(self) -> None:
        for entry in self.entries.values():
            entry.delete(0, tk.END)

    def save_record(self) -> None:
        record = {field: entry.get().strip() for field, entry in self.entries.items()}
        if not any(record.values()):
            messagebox.showwarning(
                "Registro vacío", "Debe diligenciar al menos un campo antes de guardar."
            )
            return
        self.data_manager.add_record(record)
        self.on_save()
        messagebox.showinfo("Registro", "El registro se guardó correctamente.")
        self.clear_form()


class ExternalDataTab(ttk.Frame):
    def __init__(self, master: ttk.Notebook, data_manager: DataManager):
        super().__init__(master)
        self.data_manager = data_manager
        self.external_fields: list[str] = []
        self.external_data: list[dict[str, str]] = []

        instructions = (
            "1. Seleccione el campo de llave principal de la base local.\n"
            "2. Cargue una base externa (formato CSV con encabezados).\n"
            "3. Elija los campos que desea traer desde la base externa."
        )
        ttk.Label(self, text=instructions, justify="left").pack(anchor="w", padx=10, pady=10)

        self.key_field = tk.StringVar(value=FIELDS[0])
        key_frame = ttk.Frame(self)
        key_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(key_frame, text="Campo llave local:").pack(side="left")
        ttk.OptionMenu(key_frame, self.key_field, self.key_field.get(), *FIELDS).pack(
            side="left", padx=5
        )

        load_button = ttk.Button(
            self, text="Cargar base externa (CSV)", command=self.load_external_file
        )
        load_button.pack(padx=10, pady=5, anchor="w")

        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)
        ttk.Label(list_frame, text="Campos disponibles en la base externa:").pack(anchor="w")
        self.field_listbox = tk.Listbox(
            list_frame, selectmode="multiple", exportselection=False, height=10
        )
        self.field_listbox.pack(fill="both", expand=True, pady=5)

        save_button = ttk.Button(
            self, text="Guardar configuración externa", command=self.save_configuration
        )
        save_button.pack(padx=10, pady=10, anchor="e")

    def load_external_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Seleccione archivo CSV",
            filetypes=[("CSV", "*.csv"), ("Todos los archivos", "*.*")],
        )
        if not file_path:
            return

        try:
            with open(file_path, newline="", encoding="utf-8-sig") as file:
                reader = csv.DictReader(file)
                self.external_data = [row for row in reader]
        except Exception as error:  # noqa: BLE001
            messagebox.showerror(
                "Error", f"No se pudo leer el archivo seleccionado.\n{error}"
            )
            return

        if not self.external_data:
            messagebox.showwarning(
                "Sin datos", "El archivo seleccionado no contiene registros."
            )
            return

        self.field_listbox.delete(0, tk.END)
        for column in self.external_data[0].keys():
            self.field_listbox.insert(tk.END, column)

        messagebox.showinfo(
            "Base cargada",
            "La base externa se cargó correctamente. Seleccione los campos a cruzar.",
        )

    def save_configuration(self) -> None:
        if not self.external_data:
            messagebox.showwarning(
                "Sin base externa", "Debe cargar primero una base externa."
            )
            return

        selected_indices = self.field_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(
                "Sin campos", "Seleccione al menos un campo externo."
            )
            return

        selected_fields = [self.field_listbox.get(i) for i in selected_indices]
        key_field = self.key_field.get()
        if key_field not in self.external_data[0]:
            messagebox.showwarning(
                "Campo llave no encontrado",
                "El campo llave seleccionado no existe en la base externa. Asegúrese de que el encabezado coincida.",
            )
            return

        self.data_manager.set_external_configuration(
            key_field=key_field, fields=selected_fields, data=self.external_data
        )
        messagebox.showinfo(
            "Configuración guardada",
            "Se guardó la configuración para el cruce de datos.",
        )


class DisplayConfigTab(ttk.Frame):
    def __init__(self, master: ttk.Notebook, data_manager: DataManager):
        super().__init__(master)
        self.data_manager = data_manager

        top_frame = ttk.Frame(self)
        top_frame.pack(fill="both", expand=True, padx=10, pady=10)

        selection_frame = ttk.LabelFrame(top_frame, text="Campos a mostrar en la visualización")
        selection_frame.pack(side="left", fill="both", expand=True)

        self.field_listbox = tk.Listbox(selection_frame, selectmode="multiple", exportselection=False)
        self.field_listbox.pack(fill="both", expand=True, padx=10, pady=10)

        for field in FIELDS:
            self.field_listbox.insert(tk.END, field)
            if field in self.data_manager.display_fields:
                idx = FIELDS.index(field)
                self.field_listbox.selection_set(idx)

        button_frame = ttk.Frame(selection_frame)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Button(button_frame, text="Guardar campos", command=self.save_fields).pack(side="left")

        records_frame = ttk.LabelFrame(top_frame, text="Registros guardados")
        records_frame.pack(side="left", fill="both", expand=True, padx=(10, 0))

        self.records_tree = ttk.Treeview(records_frame, columns=("Resumen",), show="headings", height=15)
        self.records_tree.heading("Resumen", text="Resumen del registro")
        self.records_tree.pack(fill="both", expand=True, padx=10, pady=10)

        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=10, pady=(0, 10))

        ttk.Button(action_frame, text="Actualizar lista", command=self.populate_records).pack(side="left")
        ttk.Button(action_frame, text="Abrir visualizador", command=self.open_viewer).pack(side="left", padx=5)

        self.populate_records()

    def populate_records(self) -> None:
        self.records_tree.delete(*self.records_tree.get_children())
        for index, record in enumerate(self.data_manager.records):
            summary_values = [
                record.get(
                    self.data_manager.display_fields[0], f"Registro {index + 1}"
                )
            ]
            self.records_tree.insert("", tk.END, iid=str(index), values=summary_values)

    def save_fields(self) -> None:
        selected_indices = self.field_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning(
                "Sin selección", "Seleccione al menos un campo para la visualización."
            )
            return
        selected_fields = [self.field_listbox.get(i) for i in selected_indices]
        self.data_manager.set_display_fields(selected_fields)
        messagebox.showinfo(
            "Campos guardados",
            "Se actualizaron los campos para la visualización.",
        )
        self.populate_records()

    def open_viewer(self) -> None:
        selected_item = self.records_tree.selection()
        if not selected_item:
            messagebox.showwarning("Sin registro", "Seleccione un registro para visualizar.")
            return
        index = int(selected_item[0])
        try:
            record = self.data_manager.records[index]
        except IndexError:
            messagebox.showerror(
                "Error", "No se pudo recuperar el registro seleccionado."
            )
            return

        RecordViewer(self, record, self.data_manager.display_fields)


class RecordViewer(tk.Toplevel):
    def __init__(self, master: tk.Widget, record: dict[str, str], fields: list[str]):
        super().__init__(master)
        self.title("Visualización del registro")
        self.geometry("600x400")

        form = ScrollableForm(self)
        form.pack(fill="both", expand=True, padx=10, pady=10)

        for index, field in enumerate(fields):
            value = record.get(field, "")
            ttk.Label(
                form.inner, text=field, style="FieldLabel.TLabel"
            ).grid(row=index, column=0, sticky="w", padx=5, pady=4)
            value_widget = ttk.Label(
                form.inner, text=value or "—", wraplength=400, justify="left"
            )
            value_widget.grid(row=index, column=1, sticky="w", padx=5, pady=4)

        ttk.Button(self, text="Cerrar", command=self.destroy).pack(pady=(0, 10))


class Application(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Herramienta de Gestión de Trámites")
        self.geometry("900x600")

        style = ttk.Style(self)
        style.configure("FieldLabel.TLabel", font=("Segoe UI", 9, "bold"))

        self.data_manager = DataManager()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        self.data_entry_tab = DataEntryTab(notebook, self.data_manager, self.refresh_records)
        self.external_data_tab = ExternalDataTab(notebook, self.data_manager)
        self.display_config_tab = DisplayConfigTab(notebook, self.data_manager)

        notebook.add(self.data_entry_tab, text="Ingreso de datos")
        notebook.add(self.external_data_tab, text="Cruce con bases externas")
        notebook.add(self.display_config_tab, text="Configuración de visualización")

    def refresh_records(self) -> None:
        self.display_config_tab.populate_records()


def main() -> None:
    app = Application()
    app.mainloop()


if __name__ == "__main__":
    main()
