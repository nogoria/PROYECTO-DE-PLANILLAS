import json
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, TYPE_CHECKING, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    import pandas as pd  # type: ignore
    _PANDAS_AVAILABLE = True
    _PANDAS_IMPORT_ERROR: Optional[ModuleNotFoundError] = None
except ModuleNotFoundError as err:
    pd = None  # type: ignore
    _PANDAS_AVAILABLE = False
    _PANDAS_IMPORT_ERROR = err

if TYPE_CHECKING:  # pragma: no cover - hints only
    import pandas as pd


CONFIG_FILE = "config.json"


@dataclass
class PlanEntry:
    plan: str
    poliza: str
    valor: float

    def to_dict(self) -> Dict[str, Any]:
        return {"PLAN": self.plan, "POLIZA": self.poliza, "VALOR": self.valor}

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "PlanEntry":
        if {"PLAN", "POLIZA", "VALOR"} <= data.keys():
            plan = str(data["PLAN"])
            poliza = str(data["POLIZA"])
            valor = float(data["VALOR"])
        else:
            plan = str(data.get("plan") or data.get("Plan") or data.get("PLAN") or "")
            poliza = str(data.get("poliza") or data.get("Poliza") or data.get("POLIZA") or "")
            valor_raw = data.get("valor") or data.get("Valor") or data.get("VALOR") or 0
            valor = float(valor_raw)
        return cls(plan=plan, poliza=poliza, valor=valor)


@dataclass
class AppConfig:
    parentescos_excluir: List[str] = field(default_factory=list)
    tipos_excluir: List[str] = field(default_factory=list)
    estados_excluir: List[str] = field(default_factory=list)
    cobro_fm: str = "No"
    t_congelada: float = 0.0
    tabla_tarifacong: str = ""
    tabla_edad: str = ""
    masculino: str = ""
    femenino: str = ""
    titulos_plan: str = ""
    planes: List[PlanEntry] = field(default_factory=list)
    tarifas: Dict[str, Dict[str, float]] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        def sort_age_key(item: Tuple[str, Dict[str, float]]) -> float:
            try:
                return float(item[0])
            except (TypeError, ValueError):
                return float("inf")

        tarifas_sorted = {
            age: {
                plan: valor
                for plan, valor in sorted(planes.items(), key=lambda entry: entry[0])
            }
            for age, planes in sorted(self.tarifas.items(), key=sort_age_key)
        }

        return {
            "Parentescos_Excluir": self.parentescos_excluir,
            "Tipos_Excluir": self.tipos_excluir,
            "Estados_Excluir": self.estados_excluir,
            "cobroFM": self.cobro_fm,
            "T_Congelada": self.t_congelada,
            "Tabla_TarifaCong": self.tabla_tarifacong,
            "Tabla_Edad": self.tabla_edad,
            "Masculino_": self.masculino,
            "Femenino_": self.femenino,
            "Titulos_Plan": self.titulos_plan,
            "planes": [plan.to_dict() for plan in self.planes],
            "tarifas": tarifas_sorted,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppConfig":
        planes_data: List[Any] = []
        if "planes" in data and isinstance(data["planes"], list):
            planes_data = data["planes"]
        elif "Planes" in data and isinstance(data["Planes"], list):
            planes_data = data["Planes"]
        planes: List[PlanEntry] = []
        for item in planes_data:
            if isinstance(item, dict):
                planes.append(PlanEntry.from_dict(item))
            else:
                # Compatibilidad con configuraciones antiguas que almacenaban solo el nombre
                planes.append(PlanEntry(plan=str(item), poliza="", valor=0.0))

        tarifas: Dict[str, Dict[str, float]] = {}
        for age, age_data in data.get("tarifas", {}).items():
            if not isinstance(age_data, dict):
                continue
            age_key = str(age)
            tarifas[age_key] = {}
            for plan, valor in age_data.items():
                try:
                    tarifas[age_key][str(plan)] = float(valor)
                except (TypeError, ValueError):
                    continue

        return cls(
            parentescos_excluir=list(data.get("Parentescos_Excluir", [])),
            tipos_excluir=list(data.get("Tipos_Excluir", [])),
            estados_excluir=list(data.get("Estados_Excluir", [])),
            cobro_fm=str(data.get("cobroFM", "No")),
            t_congelada=float(data.get("T_Congelada", 0) or 0),
            tabla_tarifacong=str(data.get("Tabla_TarifaCong", "")),
            tabla_edad=str(data.get("Tabla_Edad", "")),
            masculino=str(data.get("Masculino_", "")),
            femenino=str(data.get("Femenino_", "")),
            titulos_plan=str(data.get("Titulos_Plan", "")),
            planes=planes,
            tarifas=tarifas,
        )


class DataProcessorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Procesador de Planillas")
        self.geometry("1000x650")

        self.config_data = AppConfig()
        self.selected_file: Optional[str] = None
        self.processed_df: Optional["pd.DataFrame"] = None
        self.pandas_available = _PANDAS_AVAILABLE

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self._build_admin_tab()
        self._build_process_tab()

        self.load_configuration(initial=True)

        if not self.pandas_available:
            self._handle_missing_pandas()

    # ------------------------------------------------------------------
    # Construcción de pestañas
    # ------------------------------------------------------------------
    def _build_admin_tab(self) -> None:
        self.admin_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.admin_tab, text="Administración")

        container = ttk.Frame(self.admin_tab)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        form_frame = ttk.LabelFrame(container, text="Parámetros de configuración")
        form_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.inputs: Dict[str, tk.Variable] = {}

        def add_entry(label: str, key: str, var_type: type[tk.Variable]) -> None:
            row = ttk.Frame(form_frame)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text=label, width=20, anchor="w").pack(side="left")
            var = var_type()
            ttk.Entry(row, textvariable=var).pack(side="left", fill="x", expand=True)
            self.inputs[key] = var

        add_entry("Parentescos excluir", "Parentescos_Excluir", tk.StringVar)
        add_entry("Tipos excluir", "Tipos_Excluir", tk.StringVar)
        add_entry("Estados excluir", "Estados_Excluir", tk.StringVar)
        cobro_frame = ttk.Frame(form_frame)
        cobro_frame.pack(fill="x", pady=3)
        ttk.Label(cobro_frame, text="Cobro FM", width=20, anchor="w").pack(side="left")
        cobro_var = tk.StringVar(value="No")
        cobro_combo = ttk.Combobox(
            cobro_frame,
            textvariable=cobro_var,
            values=("Si", "No"),
            state="readonly",
            width=5,
        )
        cobro_combo.pack(side="left")
        self.inputs["cobroFM"] = cobro_var

        add_entry("T. Congelada", "T_Congelada", tk.StringVar)
        add_entry("Tabla TarifaCong", "Tabla_TarifaCong", tk.StringVar)
        add_entry("Tabla Edad", "Tabla_Edad", tk.StringVar)
        add_entry("Masculino", "Masculino_", tk.StringVar)
        add_entry("Femenino", "Femenino_", tk.StringVar)
        add_entry("Títulos Plan", "Titulos_Plan", tk.StringVar)

        # Planes y rangos
        plan_frame = ttk.LabelFrame(container, text="Planes y Rangos de Edad")
        plan_frame.pack(side="left", fill="both", expand=True, padx=(5, 0))

        list_frame = ttk.Frame(plan_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        ttk.Label(list_frame, text="Planes").pack(anchor="w")
        plan_columns = ("plan", "poliza", "valor")
        self.plan_tree = ttk.Treeview(
            list_frame, columns=plan_columns, show="headings", height=8
        )
        for column, heading in zip(plan_columns, ("Plan", "Póliza", "Valor")):
            anchor = "center" if column != "plan" else "w"
            self.plan_tree.heading(column, text=heading)
            self.plan_tree.column(column, width=110, anchor=anchor)
        self.plan_tree.pack(side="left", fill="both", expand=True, pady=5)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.plan_tree.yview)
        self.plan_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self.plan_tree.bind("<Double-1>", self.edit_plan)

        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Agregar Plan", command=self.add_plan).pack(
            side="left", expand=True, fill="x", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="Eliminar Plan", command=self.remove_plan).pack(
            side="left", expand=True, fill="x", padx=(5, 0)
        )

        tariff_frame = ttk.LabelFrame(plan_frame, text="Tarifas por Edad y Plan")
        tariff_frame.pack(fill="both", expand=True)
        self.tariff_columns: List[str] = []
        self.selected_tariff_column: Optional[str] = None
        self.tariff_tree = ttk.Treeview(tariff_frame, show="headings", height=8)
        self.tariff_tree.pack(fill="both", expand=True, pady=5)
        self.tariff_tree.bind("<Double-1>", self._on_tariff_double_click)
        self.tariff_tree.bind("<Button-1>", self._on_tariff_click)
        tariff_scroll = ttk.Scrollbar(
            tariff_frame, orient="vertical", command=self.tariff_tree.yview
        )
        self.tariff_tree.configure(yscrollcommand=tariff_scroll.set)
        tariff_scroll.pack(side="right", fill="y")

        tariff_btn_frame = ttk.Frame(tariff_frame)
        tariff_btn_frame.pack(fill="x", pady=5)
        ttk.Button(
            tariff_btn_frame,
            text="Agregar Rango de Edad",
            command=self._add_tariff_age_row,
        ).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ttk.Button(
            tariff_btn_frame,
            text="Eliminar Rango",
            command=self._remove_tariff_age_row,
        ).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(
            tariff_btn_frame,
            text="Agregar Plan",
            command=self._add_tariff_plan_column,
        ).pack(side="left", expand=True, fill="x", padx=5)
        ttk.Button(
            tariff_btn_frame,
            text="Eliminar Plan",
            command=self._remove_tariff_plan_column,
        ).pack(side="left", expand=True, fill="x", padx=(5, 0))

        self.tariff_edit_entry: Optional[tk.Entry] = None
        self._refresh_tariff_tree()

        action_frame = ttk.Frame(self.admin_tab)
        action_frame.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(
            action_frame, text="Guardar configuración", command=self.save_configuration
        ).pack(side="left", padx=5)
        ttk.Button(
            action_frame, text="Cargar configuración", command=self.load_configuration
        ).pack(side="left", padx=5)

    def _build_process_tab(self) -> None:
        self.process_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.process_tab, text="Procesar datos")

        top_frame = ttk.Frame(self.process_tab)
        top_frame.pack(fill="x", padx=10, pady=10)

        self.select_button = ttk.Button(
            top_frame, text="Seleccionar archivo Excel", command=self.select_file
        )
        self.select_button.pack(side="left")
        self.file_label = ttk.Label(top_frame, text="Ningún archivo seleccionado")
        self.file_label.pack(side="left", padx=10)

        btn_frame = ttk.Frame(self.process_tab)
        btn_frame.pack(fill="x", padx=10)
        self.process_button = ttk.Button(
            btn_frame, text="Procesar", command=self.process_data, state="disabled"
        )
        self.process_button.pack(side="left", padx=5)
        self.export_button = ttk.Button(
            btn_frame,
            text="Exportar resultado",
            command=self.export_result,
            state="disabled",
        )
        self.export_button.pack(side="left", padx=5)

        preview_frame = ttk.LabelFrame(self.process_tab, text="Vista previa del resultado")
        preview_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.preview_tree = ttk.Treeview(preview_frame, show="headings")
        self.preview_tree.pack(fill="both", expand=True)

        summary_frame = ttk.LabelFrame(self.process_tab, text="Resumen del archivo")
        summary_frame.pack(fill="x", padx=10, pady=(0, 10))
        self.summary_text = tk.Text(summary_frame, height=4, wrap="word", state="disabled")
        self.summary_text.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="Listo")
        status_bar = ttk.Label(self.process_tab, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill="x", padx=10, pady=(0, 5))

    def _handle_missing_pandas(self) -> None:
        self.select_button.config(state="disabled")
        self.process_button.config(state="disabled")
        self.export_button.config(state="disabled")
        self.status_var.set(
            "Pandas no está instalado. Instale la librería para habilitar el procesamiento."
        )

        def notify() -> None:
            message = (
                "La interfaz se cargó correctamente, pero falta la librería 'pandas'.\n"
                "Instálela con 'pip install pandas' y reinicie la aplicación para procesar archivos."
            )
            messagebox.showerror("Dependencia faltante", message)

        self.after(100, notify)

    # ------------------------------------------------------------------
    # Gestión de planes y rangos
    # ------------------------------------------------------------------
    def add_plan(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Agregar plan")
        dialog.grab_set()

        fields = {
            "Plan": tk.StringVar(),
            "Póliza": tk.StringVar(),
            "Valor": tk.StringVar(),
        }

        entries: List[ttk.Entry] = []
        for label, var in fields.items():
            frame = ttk.Frame(dialog)
            frame.pack(fill="x", padx=10, pady=5)
            ttk.Label(frame, text=label, width=10, anchor="w").pack(side="left")
            entry = ttk.Entry(frame, textvariable=var)
            entry.pack(side="left", fill="x", expand=True)
            entries.append(entry)

        if entries:
            dialog.after(50, entries[0].focus_set)

        def confirm() -> None:
            plan = fields["Plan"].get().strip()
            poliza = fields["Póliza"].get().strip()
            valor_text = fields["Valor"].get().strip()

            if not plan or not poliza or not valor_text:
                messagebox.showwarning(
                    "Datos incompletos",
                    "Complete Plan, Póliza y Valor para agregar el registro.",
                )
                return

            try:
                valor = float(valor_text)
            except ValueError:
                messagebox.showerror(
                    "Valor inválido", "Ingrese un número válido en el campo Valor."
                )
                return

            for item in self.plan_tree.get_children():
                plan_val, poliza_val, _ = self.plan_tree.item(item, "values")
                if plan_val == plan and poliza_val == poliza:
                    messagebox.showwarning(
                        "Duplicado",
                        "Ya existe un registro con el mismo Plan y Póliza.",
                    )
                    return

            self.plan_tree.insert(
                "",
                tk.END,
                values=(plan, poliza, f"{valor:.2f}"),
            )
            dialog.destroy()

        ttk.Button(dialog, text="Agregar", command=confirm).pack(pady=(0, 10))

    def remove_plan(self) -> None:
        selection = self.plan_tree.selection()
        if not selection:
            messagebox.showinfo(
                "Eliminar plan", "Seleccione un plan para eliminarlo de la tabla."
            )
            return
        for item in selection:
            self.plan_tree.delete(item)

    def edit_plan(self, event: tk.Event) -> None:
        item_id = self.plan_tree.identify_row(event.y)
        if not item_id:
            return

        current_values = self.plan_tree.item(item_id, "values")
        dialog = tk.Toplevel(self)
        dialog.title("Editar plan")
        dialog.grab_set()

        fields = {
            "Plan": tk.StringVar(value=current_values[0] if current_values else ""),
            "Póliza": tk.StringVar(value=current_values[1] if len(current_values) > 1 else ""),
            "Valor": tk.StringVar(value=current_values[2] if len(current_values) > 2 else ""),
        }

        entries: List[ttk.Entry] = []
        for label, var in fields.items():
            frame = ttk.Frame(dialog)
            frame.pack(fill="x", padx=10, pady=5)
            ttk.Label(frame, text=label, width=10, anchor="w").pack(side="left")
            entry = ttk.Entry(frame, textvariable=var)
            entry.pack(side="left", fill="x", expand=True)
            entries.append(entry)

        if entries:
            dialog.after(50, entries[0].focus_set)

        def confirm() -> None:
            plan = fields["Plan"].get().strip()
            poliza = fields["Póliza"].get().strip()
            valor_text = fields["Valor"].get().strip()

            if not plan or not poliza or not valor_text:
                messagebox.showwarning(
                    "Datos incompletos",
                    "Complete Plan, Póliza y Valor para actualizar el registro.",
                )
                return

            try:
                float(valor_text)
            except ValueError:
                messagebox.showerror(
                    "Valor inválido", "Ingrese un número válido en el campo Valor."
                )
                return

            for other_item in self.plan_tree.get_children():
                if other_item == item_id:
                    continue
                other_plan, other_poliza, _ = self.plan_tree.item(other_item, "values")
                if other_plan == plan and other_poliza == poliza:
                    messagebox.showwarning(
                        "Duplicado",
                        "Ya existe un registro con el mismo Plan y Póliza.",
                    )
                    return

            self.plan_tree.item(item_id, values=(plan, poliza, valor_text))
            dialog.destroy()

        ttk.Button(dialog, text="Actualizar", command=confirm).pack(pady=(0, 10))

    # ------------------------------------------------------------------
    # Gestión tabla de tarifas
    # ------------------------------------------------------------------
    def _refresh_tariff_tree(self) -> None:
        if self.tariff_edit_entry is not None:
            self.tariff_edit_entry.destroy()
            self.tariff_edit_entry = None
        columns = ["Edad"] + self.tariff_columns
        self.tariff_tree["columns"] = columns
        for column in columns:
            self.tariff_tree.heading(column, text=column)
            anchor = "center" if column != "Edad" else "w"
            width = 90 if column != "Edad" else 80
            self.tariff_tree.column(column, anchor=anchor, width=width, stretch=True)

        # Ensure every item has all columns
        for item in self.tariff_tree.get_children():
            current_values = list(self.tariff_tree.item(item, "values"))
            if len(current_values) > len(columns):
                current_values = current_values[: len(columns)]
            if len(current_values) < len(columns):
                current_values += [""] * (len(columns) - len(current_values))
            self.tariff_tree.item(item, values=current_values)

    def _add_tariff_age_row(self) -> None:
        value = simpledialog.askstring(
            "Agregar rango de edad",
            "Ingrese la edad base (por ejemplo 0, 60):",
            parent=self,
        )
        if value is None:
            return

        value = value.strip()
        if not value:
            messagebox.showwarning("Edad inválida", "Ingrese un valor para la edad.")
            return

        if value in {self.tariff_tree.set(item, "Edad") for item in self.tariff_tree.get_children()}:
            messagebox.showwarning(
                "Duplicado", "Ya existe un registro con la edad especificada."
            )
            return

        row_values = [value] + [""] * len(self.tariff_columns)
        self.tariff_tree.insert("", tk.END, values=row_values)

    def _remove_tariff_age_row(self) -> None:
        selection = self.tariff_tree.selection()
        if not selection:
            messagebox.showinfo(
                "Eliminar rango", "Seleccione una fila de edad para eliminarla."
            )
            return
        for item in selection:
            self.tariff_tree.delete(item)

    def _add_tariff_plan_column(self) -> None:
        plan_name = simpledialog.askstring(
            "Agregar plan", "Ingrese el identificador del plan:", parent=self
        )
        if plan_name is None:
            return
        plan_name = plan_name.strip()
        if not plan_name:
            messagebox.showwarning("Plan inválido", "Ingrese un nombre para el plan.")
            return

        normalized = plan_name.casefold()
        if any(col.casefold() == normalized for col in self.tariff_columns):
            messagebox.showwarning("Duplicado", "Ya existe una columna con ese plan.")
            return

        self.tariff_columns.append(plan_name)
        self.selected_tariff_column = None
        self._refresh_tariff_tree()

    def _remove_tariff_plan_column(self) -> None:
        if not self.tariff_columns:
            messagebox.showinfo("Eliminar plan", "No hay columnas de planes para eliminar.")
            return

        column = self.selected_tariff_column
        if column is None or column == "Edad":
            column = simpledialog.askstring(
                "Eliminar plan",
                "Indique el nombre de la columna de plan que desea eliminar:",
                parent=self,
            )
            if column is None:
                return

        column = column.strip()
        if column == "Edad" or not column:
            messagebox.showwarning(
                "Columna inválida", "Debe indicar una columna de plan para eliminar."
            )
            return

        matches = [c for c in self.tariff_columns if c.casefold() == column.casefold()]
        if not matches:
            messagebox.showerror("Plan", f"No se encontró la columna '{column}'.")
            return

        target = matches[0]
        target_index = self.tariff_columns.index(target)
        self.tariff_columns.pop(target_index)
        self.selected_tariff_column = None

        for item in self.tariff_tree.get_children():
            values = list(self.tariff_tree.item(item, "values"))
            remove_index = target_index + 1  # Offset by Edad column
            if len(values) > remove_index:
                values.pop(remove_index)
                self.tariff_tree.item(item, values=values)

        self._refresh_tariff_tree()

    def _on_tariff_click(self, event: tk.Event) -> None:
        region = self.tariff_tree.identify_region(event.x, event.y)
        if region == "heading":
            column_id = self.tariff_tree.identify_column(event.x)
            try:
                index = int(column_id.replace("#", "")) - 1
            except ValueError:
                self.selected_tariff_column = None
                return
            columns = list(self.tariff_tree["columns"])
            if 0 <= index < len(columns):
                self.selected_tariff_column = columns[index]
        else:
            self.selected_tariff_column = None

    def _on_tariff_double_click(self, event: tk.Event) -> None:
        if self.tariff_edit_entry is not None:
            self.tariff_edit_entry.destroy()
            self.tariff_edit_entry = None

        row_id = self.tariff_tree.identify_row(event.y)
        column_id = self.tariff_tree.identify_column(event.x)
        if not row_id or column_id == "#0":
            return

        x, y, width, height = self.tariff_tree.bbox(row_id, column_id)
        if width == 0 and height == 0:
            return

        column_index = int(column_id.replace("#", "")) - 1
        columns = list(self.tariff_tree["columns"])
        if column_index < 0 or column_index >= len(columns):
            return

        current_value = self.tariff_tree.set(row_id, columns[column_index])

        entry = tk.Entry(self.tariff_tree)
        entry.insert(0, current_value)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()
        entry.bind("<Return>", lambda e: self._finish_tariff_edit(row_id, columns[column_index], entry))
        entry.bind("<FocusOut>", lambda e: self._finish_tariff_edit(row_id, columns[column_index], entry))
        self.tariff_edit_entry = entry

    def _finish_tariff_edit(self, item_id: str, column: str, entry: tk.Entry) -> None:
        value = entry.get().strip()
        if column != "Edad" and value:
            try:
                float(value)
            except ValueError:
                messagebox.showerror(
                    "Valor inválido", "Ingrese un número válido para la prima."
                )
                entry.focus()
                return
        if column == "Edad":
            if not value:
                messagebox.showwarning(
                    "Edad inválida", "La edad no puede quedar vacía."
                )
                entry.focus()
                return
            for other_item in self.tariff_tree.get_children():
                if other_item == item_id:
                    continue
                if self.tariff_tree.set(other_item, "Edad") == value:
                    messagebox.showwarning(
                        "Duplicado", "Ya existe una fila con la edad indicada."
                    )
                    entry.focus()
                    return
        self.tariff_tree.set(item_id, column, value)
        entry.destroy()
        self.tariff_edit_entry = None

    # ------------------------------------------------------------------
    # Configuración
    # ------------------------------------------------------------------
    def save_configuration(self) -> None:
        config = self._gather_config_from_ui()
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as fh:
                json.dump(config.to_dict(), fh, indent=4, ensure_ascii=False)
            messagebox.showinfo("Configuración", "Configuración guardada correctamente.")
        except OSError as exc:
            messagebox.showerror("Error", f"No se pudo guardar la configuración: {exc}")

    def load_configuration(self, initial: bool = False) -> None:
        if not os.path.exists(CONFIG_FILE):
            if initial:
                self.status_var.set("Configuración por defecto cargada.")
            else:
                messagebox.showinfo(
                    "Configuración",
                    "No se encontró config.json. Se usarán valores por defecto.",
                )
            self._populate_config_ui(self.config_data)
            return

        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            self.config_data = AppConfig.from_dict(data)
            self._populate_config_ui(self.config_data)
            self.status_var.set("Configuración cargada desde config.json")
        except (json.JSONDecodeError, OSError) as exc:
            messagebox.showerror(
                "Error",
                f"No se pudo cargar la configuración: {exc}. Se usarán valores por defecto.",
            )
            self.config_data = AppConfig()
            self._populate_config_ui(self.config_data)

    def _gather_config_from_ui(self) -> AppConfig:
        def split_values(text: str) -> List[str]:
            return [item.strip() for item in text.split(",") if item.strip()]

        parentescos = split_values(self.inputs["Parentescos_Excluir"].get())
        tipos = split_values(self.inputs["Tipos_Excluir"].get())
        estados = split_values(self.inputs["Estados_Excluir"].get())

        try:
            t_congelada = float(self.inputs["T_Congelada"].get())
        except ValueError:
            t_congelada = 0.0

        planes: List[PlanEntry] = []
        for item in self.plan_tree.get_children():
            plan, poliza, valor = self.plan_tree.item(item, "values")
            try:
                valor_float = float(valor)
            except (TypeError, ValueError):
                valor_float = 0.0
            planes.append(PlanEntry(plan=str(plan), poliza=str(poliza), valor=valor_float))

        tarifas: Dict[str, Dict[str, float]] = {}
        columns = list(self.tariff_tree["columns"])
        for item in self.tariff_tree.get_children():
            row_values = list(self.tariff_tree.item(item, "values"))
            if not row_values:
                continue
            edad = str(row_values[0]).strip()
            if not edad:
                continue
            tarifas.setdefault(edad, {})
            for index, column in enumerate(columns[1:], start=1):
                if index >= len(row_values):
                    continue
                value_text = str(row_values[index]).strip()
                if not value_text:
                    continue
                try:
                    value_float = float(value_text)
                except ValueError:
                    continue
                tarifas[edad][column] = value_float

        config = AppConfig(
            parentescos_excluir=parentescos,
            tipos_excluir=tipos,
            estados_excluir=estados,
            cobro_fm=self.inputs["cobroFM"].get(),
            t_congelada=t_congelada,
            tabla_tarifacong=self.inputs["Tabla_TarifaCong"].get(),
            tabla_edad=self.inputs["Tabla_Edad"].get(),
            masculino=self.inputs["Masculino_"].get(),
            femenino=self.inputs["Femenino_"].get(),
            titulos_plan=self.inputs["Titulos_Plan"].get(),
            planes=planes,
            tarifas=tarifas,
        )
        self.config_data = config
        return config

    def _populate_config_ui(self, config: AppConfig) -> None:
        def join_values(values: List[str]) -> str:
            return ", ".join(values)

        self.inputs["Parentescos_Excluir"].set(join_values(config.parentescos_excluir))
        self.inputs["Tipos_Excluir"].set(join_values(config.tipos_excluir))
        self.inputs["Estados_Excluir"].set(join_values(config.estados_excluir))
        self.inputs["cobroFM"].set(config.cobro_fm)
        self.inputs["T_Congelada"].set(str(config.t_congelada))
        self.inputs["Tabla_TarifaCong"].set(config.tabla_tarifacong)
        self.inputs["Tabla_Edad"].set(config.tabla_edad)
        self.inputs["Masculino_"].set(config.masculino)
        self.inputs["Femenino_"].set(config.femenino)
        self.inputs["Titulos_Plan"].set(config.titulos_plan)

        for item in self.plan_tree.get_children():
            self.plan_tree.delete(item)
        for plan in config.planes:
            self.plan_tree.insert(
                "",
                tk.END,
                values=(plan.plan, plan.poliza, f"{plan.valor:.2f}"),
            )

        self.tariff_tree.delete(*self.tariff_tree.get_children())
        tariff_columns_set = {plan for row in config.tarifas.values() for plan in row}
        self.tariff_columns = sorted(tariff_columns_set, key=lambda value: value)
        self._refresh_tariff_tree()
        def age_sort_key(item: Tuple[str, Dict[str, float]]) -> float:
            try:
                return float(item[0])
            except (TypeError, ValueError):
                return float("inf")

        for edad, row in sorted(config.tarifas.items(), key=age_sort_key):
            values = [edad]
            for column in self.tariff_columns:
                valor = row.get(column, "")
                if valor == "":
                    values.append("")
                else:
                    values.append(f"{float(valor):.2f}")
            self.tariff_tree.insert("", tk.END, values=values)
        self._refresh_tariff_tree()

    # ------------------------------------------------------------------
    # Procesamiento de datos
    # ------------------------------------------------------------------
    def select_file(self) -> None:
        if not self.pandas_available:
            messagebox.showerror(
                "Dependencia faltante",
                "No es posible seleccionar archivos porque pandas no está instalado.",
            )
            return

        filetypes = [("Archivos de Excel", "*.xlsx *.xls"), ("Todos los archivos", "*.*")]
        path = filedialog.askopenfilename(title="Seleccionar Excel", filetypes=filetypes)
        if not path:
            return

        self.selected_file = path
        self.file_label.config(text=path)
        self.process_button.config(state="normal")
        self.status_var.set("Archivo seleccionado. Listo para procesar.")

    def process_data(self) -> None:
        if not self.pandas_available or pd is None:
            messagebox.showerror(
                "Dependencia faltante",
                "Instale la librería 'pandas' para procesar los datos.",
            )
            return

        if not self.selected_file:
            messagebox.showwarning("Procesar", "Seleccione un archivo antes de procesar.")
            return

        config = self._gather_config_from_ui()

        try:
            df = pd.read_excel(self.selected_file, sheet_name="Base")
        except FileNotFoundError:
            messagebox.showerror("Archivo", "El archivo seleccionado no existe.")
            return
        except ValueError as exc:
            messagebox.showerror("Hoja inválida", f"No se encontró la hoja 'Base': {exc}")
            return
        except Exception as exc:
            messagebox.showerror("Lectura", f"No fue posible leer el archivo: {exc}")
            return

        if df.empty:
            messagebox.showinfo("Procesar", "La hoja 'Base' no contiene datos.")
            return

        self._update_summary(df)

        self.status_var.set("Aplicando filtros y cálculos...")
        self.update_idletasks()

        processed = self._apply_business_logic(df, config)
        self.processed_df = processed
        self._populate_preview(processed)

        self.export_button.config(state="normal")
        self.status_var.set("Procesamiento completado. Puede exportar el resultado.")

    def _apply_business_logic(self, df: "pd.DataFrame", config: AppConfig) -> "pd.DataFrame":
        result = df.copy()

        # Filtrar por exclusiones si las columnas existen
        filters = [
            ("Parentesco", config.parentescos_excluir),
            ("Tipo", config.tipos_excluir),
            ("Estado", config.estados_excluir),
        ]
        for column, exclusions in filters:
            if exclusions and column in result.columns:
                result = result[~result[column].isin(exclusions)]

        def normalize_string(value: Any) -> str:
            if value is None:
                return ""
            if pd is not None and pd.isna(value):  # type: ignore[attr-defined]
                return ""
            return str(value).strip().casefold()

        def find_column(target: str) -> Optional[str]:
            normalized_target = target
            for column in result.columns:
                normalized = (
                    column.lower()
                    .replace("á", "a")
                    .replace("é", "e")
                    .replace("í", "i")
                    .replace("ó", "o")
                    .replace("ú", "u")
                )
                if normalized == normalized_target:
                    return column
            return None

        plan_column = find_column("plan")
        poliza_column = find_column("poliza")
        edad_column = find_column("edad")

        if config.planes and plan_column:
            plan_values = {normalize_string(entry.plan) for entry in config.planes if entry.plan}
            result["Plan_Valido"] = result[plan_column].apply(
                lambda value: normalize_string(value) in plan_values
            )
        elif plan_column:
            result["Plan_Valido"] = True

        plan_mapping = {
            (normalize_string(entry.plan), normalize_string(entry.poliza)): entry.valor
            for entry in config.planes
            if entry.plan and entry.poliza
        }

        if "Descuento POS" in result.columns:
            descuento_pos: List[Optional[float]] = list(result["Descuento POS"])
        else:
            descuento_pos = [None] * len(result)

        if plan_mapping and plan_column and poliza_column:
            for idx, row in result.iterrows():
                key = (
                    normalize_string(row.get(plan_column)),
                    normalize_string(row.get(poliza_column)),
                )
                valor = plan_mapping.get(key)
                if valor is not None:
                    descuento_pos[idx] = valor

        # Calcular Prima Neta
        tarifas = {
            float(age_key): {plan: float(valor) for plan, valor in planes.items()}
            for age_key, planes in config.tarifas.items()
            if _is_number(age_key) and isinstance(planes, dict)
        }

        if "Prima Neta" in result.columns:
            prima_neta: List[Optional[float]] = list(result["Prima Neta"])
        else:
            prima_neta = [None] * len(result)

        if tarifas and plan_column and edad_column:
            sorted_ages = sorted(tarifas.keys())

            def obtener_tarifa(edad: Any, plan_val: Any) -> Optional[float]:
                try:
                    edad_float = float(edad)
                except (TypeError, ValueError):
                    return None

                plan_normalized = normalize_string(plan_val)
                applicable_age = None
                for age_threshold in sorted_ages:
                    if edad_float >= age_threshold:
                        applicable_age = age_threshold
                if applicable_age is None:
                    applicable_age = sorted_ages[0]

                row_tarifas = tarifas.get(applicable_age, {})
                for plan, valor in row_tarifas.items():
                    if normalize_string(plan) == plan_normalized:
                        return valor
                return None

            for idx, row in result.iterrows():
                valor = obtener_tarifa(row.get(edad_column), row.get(plan_column))
                if valor is not None:
                    prima_neta[idx] = valor

        if "Prima Neta" in result.columns:
            result.loc[:, "Prima Neta"] = prima_neta
        else:
            result["Prima Neta"] = prima_neta

        if "Descuento POS" in result.columns:
            result.loc[:, "Descuento POS"] = descuento_pos
        else:
            result["Descuento POS"] = descuento_pos

        result.reset_index(drop=True, inplace=True)
        return result

    def _update_summary(self, df: "pd.DataFrame") -> None:
        info = [
            f"Filas: {len(df)}",
            f"Columnas: {len(df.columns)}",
            "Lista de columnas: " + ", ".join(map(str, df.columns)),
        ]
        self.summary_text.config(state="normal")
        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert(tk.END, "\n".join(info))
        self.summary_text.config(state="disabled")

    def _populate_preview(self, df: "pd.DataFrame") -> None:
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = list(df.columns)
        for col in df.columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=120, anchor="center")

        max_rows = 100
        for _, row in df.head(max_rows).iterrows():
            values = []
            for col in df.columns:
                value = row[col]
                if pd is not None:
                    values.append(value if pd.notna(value) else "")
                else:
                    values.append("" if value is None else value)
            self.preview_tree.insert("", tk.END, values=values)

    def export_result(self) -> None:
        if self.processed_df is None:
            messagebox.showwarning("Exportar", "No hay datos para exportar.")
            return

        if not self.pandas_available or pd is None:
            messagebox.showerror(
                "Dependencia faltante",
                "Instale la librería 'pandas' para exportar resultados.",
            )
            return

        path = filedialog.asksaveasfilename(
            title="Guardar resultado",
            defaultextension=".xlsx",
            filetypes=[("Archivo de Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not path:
            return

        try:
            self.processed_df.to_excel(path, index=False)
            messagebox.showinfo("Exportar", "Archivo guardado correctamente.")
        except Exception as exc:
            messagebox.showerror("Exportar", f"No fue posible guardar el archivo: {exc}")


def _is_number(value: Any) -> bool:
    try:
        float(value)
    except (TypeError, ValueError):
        return False
    return True


if __name__ == "__main__":
    app = DataProcessorApp()
    app.mainloop()
