import json
import os
import unicodedata
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


def _parse_numeric(value: Any) -> Optional[float]:
    """Convert diverse numeric string formats into floats when possible."""

    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = text.replace(" ", "")

    candidates = [text]
    if "," in text and "." in text:
        candidates.append(text.replace(".", "").replace(",", "."))
    if "," in text:
        candidates.append(text.replace(",", "."))
    if text.count(".") > 1:
        candidates.append(text.replace(".", ""))

    for candidate in candidates:
        try:
            return float(candidate)
        except ValueError:
            continue

    return None


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

        tarifas_frame = ttk.LabelFrame(plan_frame, text="Tarifas por Edad y Plan")
        tarifas_frame.pack(fill="both", expand=True)

        self.tarifas_tree = ttk.Treeview(tarifas_frame, columns=("Edad",), show="headings")
        self.tarifas_tree.heading("Edad", text="Edad")
        self.tarifas_tree.column("Edad", width=100, anchor="center")
        self.tarifas_tree.pack(fill="both", expand=True)

        self.tarifas: Dict[str, Dict[str, Any]] = {}

        tarifas_btns = ttk.Frame(tarifas_frame)
        tarifas_btns.pack(fill="x", pady=(5, 5))

        ttk.Button(
            tarifas_btns,
            text="Agregar Rango de Edad",
            command=self.agregar_rango_edad,
        ).pack(side="left", padx=5)
        ttk.Button(
            tarifas_btns,
            text="Eliminar Rango",
            command=self.eliminar_rango_edad,
        ).pack(side="left", padx=5)
        ttk.Button(
            tarifas_btns,
            text="Agregar Plan",
            command=self.agregar_plan_a_tarifas,
        ).pack(side="left", padx=5)
        ttk.Button(
            tarifas_btns,
            text="Eliminar Plan",
            command=self.eliminar_plan_de_tarifas,
        ).pack(side="left", padx=5)

        self.tarifas_tree.bind("<Double-1>", self.editar_celda_tarifa)

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
    def agregar_rango_edad(self) -> None:
        rango = simpledialog.askstring(
            "Nuevo rango", "Ingrese el rango de edad (ej: 0,59):", parent=self
        )
        if not rango:
            return

        rango = rango.strip()
        if rango in self.tarifas:
            messagebox.showwarning("Aviso", "Este rango ya existe.")
            return

        self.tarifas[rango] = {}
        self.tarifas_tree.insert("", "end", values=(rango,))
        self.actualizar_columnas_tarifas()

    def eliminar_rango_edad(self) -> None:
        selected = self.tarifas_tree.selection()
        if not selected:
            return
        item_id = selected[0]
        valores = self.tarifas_tree.item(item_id, "values")
        rango = valores[0] if valores else ""
        if rango in self.tarifas:
            del self.tarifas[rango]
        self.tarifas_tree.delete(*selected)
        self.actualizar_columnas_tarifas()

    def agregar_plan_a_tarifas(self) -> None:
        plan = simpledialog.askstring(
            "Nuevo plan", "Ingrese el número del plan:", parent=self
        )
        if not plan:
            return

        plan = plan.strip()
        if plan in self.tarifas_tree["columns"]:
            messagebox.showwarning("Aviso", "Este plan ya existe.")
            return

        cols = list(self.tarifas_tree["columns"]) + [plan]
        self.tarifas_tree["columns"] = cols
        self.tarifas_tree.heading(plan, text=plan)
        self.tarifas_tree.column(plan, width=100, anchor="center")

        for rango in self.tarifas:
            self.tarifas[rango][plan] = self.tarifas[rango].get(plan, "")

        self.actualizar_columnas_tarifas()

    def eliminar_plan_de_tarifas(self) -> None:
        plan = simpledialog.askstring(
            "Eliminar plan", "Ingrese el número del plan a eliminar:", parent=self
        )
        if not plan:
            return
        if plan == "Edad":
            messagebox.showwarning("Aviso", "La columna Edad no puede eliminarse.")
            return
        if plan not in self.tarifas_tree["columns"]:
            return

        for rango in self.tarifas:
            self.tarifas[rango].pop(plan, None)

        nuevas = [c for c in self.tarifas_tree["columns"] if c != plan]
        self.tarifas_tree["columns"] = nuevas
        self.actualizar_columnas_tarifas()

    def editar_celda_tarifa(self, event: tk.Event) -> None:
        region = self.tarifas_tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        col = self.tarifas_tree.identify_column(event.x)
        row = self.tarifas_tree.identify_row(event.y)
        if not row or col == "#1":
            return

        try:
            col_index = int(col.replace("#", "")) - 1
        except ValueError:
            return

        columnas = self.tarifas_tree["columns"]
        if col_index < 0 or col_index >= len(columnas):
            return

        col_name = columnas[col_index]
        rango = self.tarifas_tree.item(row, "values")[0]

        bbox = self.tarifas_tree.bbox(row, col)
        if not bbox:
            return
        x, y, width, height = bbox

        entry = ttk.Entry(self.tarifas_tree)
        valor_actual = self.tarifas.get(rango, {}).get(col_name, "")
        entry.insert(0, str(valor_actual))
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus()

        def guardar_valor(event: Optional[tk.Event] = None) -> None:
            nuevo_valor = entry.get().strip()
            entry.destroy()

            if nuevo_valor == "":
                valor_final: Any = ""
            else:
                try:
                    valor_final = float(nuevo_valor.replace(",", "."))
                except ValueError:
                    valor_final = nuevo_valor

            self.tarifas.setdefault(rango, {})[col_name] = valor_final
            self.actualizar_columnas_tarifas()

        entry.bind("<Return>", guardar_valor)
        entry.bind("<FocusOut>", guardar_valor)

    def actualizar_columnas_tarifas(self) -> None:
        columnas = list(self.tarifas_tree["columns"])
        if "Edad" not in columnas:
            columnas = ["Edad"] + [col for col in columnas if col != "Edad"]
            self.tarifas_tree["columns"] = columnas

        for column in columnas:
            texto = "Edad" if column == "Edad" else column
            self.tarifas_tree.heading(column, text=texto)
            self.tarifas_tree.column(column, width=100, anchor="center")

        for item in self.tarifas_tree.get_children():
            self.tarifas_tree.delete(item)

        for rango, valores in self.tarifas.items():
            fila = [rango]
            for plan in columnas[1:]:
                valor = valores.get(plan, "")
                if isinstance(valor, float) and valor.is_integer():
                    fila.append(str(int(valor)))
                else:
                    fila.append(str(valor) if valor != "" else "")
            self.tarifas_tree.insert("", "end", values=fila)

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

        t_congelada_value = _parse_numeric(self.inputs["T_Congelada"].get())
        t_congelada = t_congelada_value if t_congelada_value is not None else 0.0

        planes: List[PlanEntry] = []
        for item in self.plan_tree.get_children():
            plan, poliza, valor = self.plan_tree.item(item, "values")
            valor_float = _parse_numeric(valor)
            if valor_float is None:
                valor_float = 0.0
            planes.append(PlanEntry(plan=str(plan), poliza=str(poliza), valor=valor_float))

        tarifas: Dict[str, Dict[str, float]] = {}
        for rango, planes_dict in self.tarifas.items():
            rango_key = str(rango).strip()
            if not rango_key:
                continue
            tarifas[rango_key] = {}
            for plan, valor in planes_dict.items():
                if valor in ("", None):
                    continue
                valor_float = _parse_numeric(valor)
                if valor_float is None:
                    continue
                tarifas[rango_key][str(plan)] = valor_float

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

        self.tarifas.clear()
        for edad, row in config.tarifas.items():
            self.tarifas[str(edad)] = {str(plan): valor for plan, valor in row.items()}

        plan_columns = sorted(
            {plan for tarifas in self.tarifas.values() for plan in tarifas.keys()}
        )
        columnas = ["Edad"] + plan_columns
        self.tarifas_tree["columns"] = columnas
        for column in columnas:
            texto = "Edad" if column == "Edad" else column
            self.tarifas_tree.heading(column, text=texto)
            self.tarifas_tree.column(column, width=100, anchor="center")

        self.actualizar_columnas_tarifas()

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

        result = self.filtrar_beneficiarios(result)

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

        result = self.asignar_prima_neta(result)

        result = self.aplicar_politica_beneficios(result)

        if "Descuento POS" in result.columns:
            result.loc[:, "Descuento POS"] = descuento_pos
        else:
            result["Descuento POS"] = descuento_pos

        result.reset_index(drop=True, inplace=True)
        return result

    def filtrar_beneficiarios(self, df: "pd.DataFrame") -> "pd.DataFrame":
        """Marca beneficiarios elegibles respetando transiciones de estado civil."""

        if pd is None or df.empty:
            if "Elegible_Beneficio" not in df.columns:
                df["Elegible_Beneficio"] = False
            if "Transicion_Estado_Civil" not in df.columns:
                df["Transicion_Estado_Civil"] = ""
            return df

        work_df = df
        if "Elegible_Beneficio" not in work_df.columns:
            work_df["Elegible_Beneficio"] = False
        else:
            work_df["Elegible_Beneficio"] = work_df["Elegible_Beneficio"].fillna(False)

        if "Transicion_Estado_Civil" not in work_df.columns:
            work_df["Transicion_Estado_Civil"] = ""
        else:
            work_df["Transicion_Estado_Civil"] = work_df["Transicion_Estado_Civil"].fillna("")

        def normalize_column_name(name: Any) -> str:
            base = str(name).strip().casefold()
            return base.replace(" ", "_")

        column_map: Dict[str, Optional[str]] = {
            "identificacion_titular": None,
            "estado_civil": None,
            "parentesco": None,
        }

        for column in work_df.columns:
            normalized = normalize_column_name(column)
            if normalized in column_map and column_map[normalized] is None:
                column_map[normalized] = column

        titular_column = column_map["identificacion_titular"]
        estado_column = column_map["estado_civil"]
        parentesco_column = column_map["parentesco"]

        if not all([titular_column, estado_column, parentesco_column]):
            return work_df

        for titular, grupo in work_df.groupby(titular_column):
            try:
                estado_valor = grupo.iloc[0][estado_column]
                if pd.isna(estado_valor):  # type: ignore[attr-defined]
                    estado_civil = ""
                else:
                    estado_civil = str(estado_valor).strip().lower()

                parentescos_series = grupo[parentesco_column]
                if pd.isna(parentescos_series).all():  # type: ignore[attr-defined]
                    parentescos_series = parentescos_series.fillna("")

                tiene_padres = parentescos_series.isin(["Padre", "Madre"]).any()

                if (
                    ("casado" in estado_civil or "compañero" in estado_civil)
                    and tiene_padres
                ):
                    work_df.loc[grupo.index, "Transicion_Estado_Civil"] = "Soltero→Casado"

                if "casado" in estado_civil or "compañero" in estado_civil:
                    parentescos_validos = ["Cónyuge", "Compañero(a)", "Hijo", "Hija"]
                    parentescos_excluidos = ["Padre", "Madre"]
                else:
                    parentescos_validos = ["Padre", "Madre", "Hijo", "Hija"]
                    parentescos_excluidos = []

                beneficiarios_validos = grupo[
                    grupo[parentesco_column].isin(parentescos_validos)
                ].copy()

                if len(beneficiarios_validos) > 3:
                    if "casado" in estado_civil:
                        prioridad = ["Cónyuge", "Compañero(a)", "Hijo", "Hija"]
                    else:
                        prioridad = ["Padre", "Madre", "Hijo", "Hija"]

                    beneficiarios_validos["prioridad"] = beneficiarios_validos[
                        parentesco_column
                    ].apply(lambda p: prioridad.index(p) if p in prioridad else 99)
                    beneficiarios_validos = beneficiarios_validos.sort_values("prioridad").head(3)

                work_df.loc[beneficiarios_validos.index, "Elegible_Beneficio"] = True

                if parentescos_excluidos:
                    padres = grupo[grupo[parentesco_column].isin(parentescos_excluidos)].index
                    work_df.loc[padres, "Elegible_Beneficio"] = False

            except Exception as exc:
                print(f"Error procesando titular {titular}: {exc}")

        return work_df

    def asignar_prima_neta(self, df: "pd.DataFrame") -> "pd.DataFrame":
        """
        Cruza el DataFrame con las tarifas definidas por plan y rango de edad (decimales incluidos),
        y asigna el valor correspondiente en la columna 'Prima Neta'.
        """

        if pd is None or df.empty:
            return df

        tarifas_fuente: Dict[str, Dict[str, Any]] = {}
        if self.config_data.tarifas:
            tarifas_fuente = self.config_data.tarifas
        elif hasattr(self, "tarifas"):
            tarifas_fuente = self.tarifas  # type: ignore[assignment]

        if not tarifas_fuente:
            return df

        plan_column: Optional[str] = None
        edad_column: Optional[str] = None
        for column in df.columns:
            normalized = str(column).strip().lower()
            if normalized == "plan" and plan_column is None:
                plan_column = column
            elif normalized == "edad" and edad_column is None:
                edad_column = column

        if plan_column is None or edad_column is None:
            return df

        parsed_tarifas: List[Tuple[float, float, Dict[str, float]]] = []
        for rango, valores in tarifas_fuente.items():
            if not isinstance(valores, dict):
                continue
            try:
                limites = [parte.strip() for parte in str(rango).split(",")]
                if len(limites) != 2:
                    continue
                edad_min_val = _parse_numeric(limites[0])
                edad_max_val = _parse_numeric(limites[1])
                if edad_min_val is None or edad_max_val is None:
                    continue
                planes_normalizados: Dict[str, float] = {}
                for plan_clave, valor in valores.items():
                    plan_nombre = str(plan_clave).strip()
                    if not plan_nombre:
                        continue
                    valor_float = _parse_numeric(valor)
                    if valor_float is None:
                        continue
                    planes_normalizados[plan_nombre.casefold()] = valor_float
                if planes_normalizados:
                    parsed_tarifas.append((edad_min_val, edad_max_val, planes_normalizados))
            except Exception:
                continue

        if not parsed_tarifas:
            return df

        if "Prima Neta" not in df.columns:
            df["Prima Neta"] = 0

        for i, row in df.iterrows():
            try:
                plan_valor = row.get(plan_column)
                if pd.isna(plan_valor):  # type: ignore[attr-defined]
                    plan = ""
                else:
                    plan = str(plan_valor).strip()
                if not plan:
                    df.at[i, "Prima Neta"] = 0
                    continue

                edad_valor = row.get(edad_column)
                edad_numero = _parse_numeric(edad_valor)
                if edad_numero is None:
                    df.at[i, "Prima Neta"] = 0
                    continue

                prima: Optional[float] = None
                for edad_min, edad_max, planes in parsed_tarifas:
                    if edad_min <= edad_numero <= edad_max:
                        prima = planes.get(plan.casefold())
                        if prima is not None:
                            break

                if prima is not None:
                    df.at[i, "Prima Neta"] = prima
                else:
                    df.at[i, "Prima Neta"] = 0

            except Exception as exc:
                print(f"Error al procesar fila {i}: {exc}")
                df.at[i, "Prima Neta"] = 0

        return df



    def aplicar_politica_beneficios(self, df: "pd.DataFrame") -> "pd.DataFrame":
        """Aplica las políticas de beneficio familiar con mensajes detallados."""

        if df.empty or pd is None:
            defaults = {
                "Aplica_Beneficio": False,
                "Motivo_No_Beneficio": "",
                "Transicion_Soltero_Casado": False,
                "Porcentaje_Beneficio": 0.0,
            }
            for columna, valor in defaults.items():
                if columna not in df.columns:
                    df[columna] = valor
                else:
                    df[columna] = df[columna].fillna(valor)
            return df

        tabla_valores = pd.DataFrame([
            {"TIPO POLIZA": "Salud Sura Clasica", "PLAN": "266", "VALOR": 89000},
            {"TIPO POLIZA": "Salud Sura Clasica", "PLAN": "267", "VALOR": 89000},
            {"TIPO POLIZA": "Sura Evoluciona", "PLAN": "817", "VALOR": 71000},
            {"TIPO POLIZA": "Salud Sura Global", "PLAN": "307", "VALOR": 89000},
            {"TIPO POLIZA": "SALUD PARA TODOS", "PLAN": "13", "VALOR": 57000},
            {"TIPO POLIZA": "SALUD PARA TODOS", "PLAN": "11", "VALOR": 57000},
            {"TIPO POLIZA": "SALUD PARA TODOS", "PLAN": "12", "VALOR": 57000},
        ])

        valores_dict = {
            (str(row["TIPO POLIZA"]).strip().lower(), str(row["PLAN"]).strip()): row["VALOR"]
            for _, row in tabla_valores.iterrows()
        }

        df["Aplica_Beneficio"] = False
        df["Motivo_No_Beneficio"] = ""
        df["Transicion_Soltero_Casado"] = False
        df["Porcentaje_Beneficio"] = 0.0

        if "Identificacion_Titular" not in df.columns or "Parentesco" not in df.columns:
            return df

        for titular_id, grupo in df.groupby("Identificacion_Titular"):
            grupo_texto = str(grupo.iloc[0].get("Grupo", "")).lower()
            parentescos = grupo["Parentesco"].astype(str).str.lower().tolist()

            tiene_conyuge = any("cónyuge" in p or "conyuge" in p or "compañero" in p for p in parentescos)

            if "soltero" in grupo_texto and tiene_conyuge:
                df.loc[grupo.index, "Transicion_Soltero_Casado"] = True
                modo = "Transición Soltero → Casado"
                prioridad = ["titular", "cónyuge", "conyuge", "compañero", "hijo", "padre", "madre"]
            elif any(palabra in grupo_texto for palabra in ["casado", "cónyuge", "conyuge", "compañero"]):
                modo = "Casado"
                prioridad = ["titular", "cónyuge", "conyuge", "compañero", "hijo"]
            else:
                modo = "Soltero"
                prioridad = ["titular", "padre", "madre", "hijo"]

            max_beneficiarios = 3
            aplican: List[int] = []

            for p in prioridad:
                subset = grupo[grupo["Parentesco"].str.lower().str.contains(p)]
                for idx in subset.index:
                    if len(aplican) < max_beneficiarios:
                        aplican.append(idx)
                    else:
                        df.at[idx, "Motivo_No_Beneficio"] = (
                            f"💔 Excede el máximo de {max_beneficiarios} beneficiarios del grupo '{modo}'."
                        )

            for idx, row in grupo.iterrows():
                parentesco = str(row.get("Parentesco", "")).lower()
                tipo_poliza = str(row.get("TIPO POLIZA", "")).strip().lower()
                plan = str(row.get("PLAN", "")).strip()

                valor_beneficio = valores_dict.get((tipo_poliza, plan), 0)

                if idx in aplican and valor_beneficio > 0:
                    df.at[idx, "Aplica_Beneficio"] = True
                    df.at[idx, "Porcentaje_Beneficio"] = valor_beneficio
                    df.at[idx, "Motivo_No_Beneficio"] = f"✅ Aplica según grupo '{modo}' y plan {plan}."
                else:
                    if df.at[idx, "Motivo_No_Beneficio"] == "":
                        if valor_beneficio == 0:
                            df.at[idx, "Motivo_No_Beneficio"] = (
                                "⚠️ No existe valor configurado en la tabla de beneficios para este plan/tipo de póliza."
                            )
                        elif "casado" in grupo_texto and any(p in parentesco for p in ["padre", "madre"]):
                            df.at[idx, "Motivo_No_Beneficio"] = (
                                "💔 No aplica porque el grupo es Casado y el parentesco es Padre/Madre."
                            )
                        elif "soltero" in grupo_texto and any(p in parentesco for p in ["cónyuge", "conyuge", "compañero"]):
                            df.at[idx, "Motivo_No_Beneficio"] = (
                                "💔 Grupo indica Soltero, pero se detectó Cónyuge/Compañero(a). Se considera transición."
                            )
                        else:
                            df.at[idx, "Motivo_No_Beneficio"] = (
                                f"💔 No aplica por regla del grupo '{modo}'. Parentesco fuera de prioridad."
                            )

        return df

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
if __name__ == "__main__":
    app = DataProcessorApp()
    app.mainloop()
