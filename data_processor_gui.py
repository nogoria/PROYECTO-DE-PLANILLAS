import json
import os
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, TYPE_CHECKING

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

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
class AgeRange:
    edad_min: float
    edad_max: float
    valor: float

    def to_dict(self) -> Dict[str, float]:
        return {
            "edad_min": self.edad_min,
            "edad_max": self.edad_max,
            "valor": self.valor,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AgeRange":
        return cls(float(data["edad_min"]), float(data["edad_max"]), float(data["valor"]))


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
    planes: List[str] = field(default_factory=list)
    rangos: List[AgeRange] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
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
            "Planes": self.planes,
            "Rangos": [r.to_dict() for r in self.rangos],
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "AppConfig":
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
            planes=list(data.get("Planes", [])),
            rangos=[AgeRange.from_dict(r) for r in data.get("Rangos", [])],
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
        self.plan_listbox = tk.Listbox(list_frame, height=8)
        self.plan_listbox.pack(fill="x", pady=5)

        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill="x", pady=5)
        ttk.Button(btn_frame, text="Agregar Plan", command=self.add_plan).pack(
            side="left", expand=True, fill="x", padx=(0, 5)
        )
        ttk.Button(btn_frame, text="Eliminar Plan", command=self.remove_plan).pack(
            side="left", expand=True, fill="x", padx=(5, 0)
        )

        range_frame = ttk.Frame(plan_frame)
        range_frame.pack(fill="both", expand=True)
        ttk.Label(range_frame, text="Rangos de Edad").pack(anchor="w")
        columns = ("edad_min", "edad_max", "valor")
        self.range_tree = ttk.Treeview(
            range_frame, columns=columns, show="headings", height=8
        )
        for col, text in zip(columns, ("Edad mínima", "Edad máxima", "Valor")):
            self.range_tree.heading(col, text=text)
            self.range_tree.column(col, width=90, anchor="center")
        self.range_tree.pack(fill="both", expand=True, pady=5)

        range_btn_frame = ttk.Frame(range_frame)
        range_btn_frame.pack(fill="x", pady=5)
        ttk.Button(
            range_btn_frame, text="Agregar Rango de Edad", command=self.add_range
        ).pack(side="left", expand=True, fill="x", padx=(0, 5))
        ttk.Button(
            range_btn_frame, text="Eliminar Rango", command=self.remove_range
        ).pack(side="left", expand=True, fill="x", padx=(5, 0))

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

        ttk.Label(dialog, text="Nombre del plan:").pack(padx=10, pady=(10, 5))
        name_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name_var)
        entry.pack(padx=10, pady=5)
        entry.focus()

        def confirm() -> None:
            value = name_var.get().strip()
            if not value:
                messagebox.showwarning("Entrada inválida", "Ingrese un nombre de plan válido.")
                return
            current = list(self.plan_listbox.get(0, tk.END))
            if value in current:
                messagebox.showwarning("Duplicado", "El plan ya existe en la lista.")
                return
            self.plan_listbox.insert(tk.END, value)
            dialog.destroy()

        ttk.Button(dialog, text="Agregar", command=confirm).pack(pady=(0, 10))

    def remove_plan(self) -> None:
        selection = self.plan_listbox.curselection()
        if not selection:
            messagebox.showinfo("Eliminar plan", "Seleccione un plan para eliminarlo.")
            return
        self.plan_listbox.delete(selection[0])

    def add_range(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("Agregar rango de edad")
        dialog.grab_set()

        fields = {
            "Edad mínima": tk.StringVar(),
            "Edad máxima": tk.StringVar(),
            "Valor": tk.StringVar(),
        }

        for label, var in fields.items():
            frame = ttk.Frame(dialog)
            frame.pack(fill="x", padx=10, pady=5)
            ttk.Label(frame, text=label, width=15).pack(side="left")
            ttk.Entry(frame, textvariable=var).pack(side="left", fill="x", expand=True)

        def confirm() -> None:
            try:
                edad_min = float(fields["Edad mínima"].get())
                edad_max = float(fields["Edad máxima"].get())
                valor = float(fields["Valor"].get())
            except ValueError:
                messagebox.showerror(
                    "Datos inválidos", "Ingrese valores numéricos para el rango y el valor."
                )
                return

            if edad_min > edad_max:
                messagebox.showerror(
                    "Rango inválido", "La edad mínima no puede ser mayor que la máxima."
                )
                return

            self.range_tree.insert(
                "",
                tk.END,
                values=(
                    f"{edad_min:.0f}" if edad_min.is_integer() else f"{edad_min}",
                    f"{edad_max:.0f}" if edad_max.is_integer() else f"{edad_max}",
                    f"{valor:.2f}",
                ),
            )
            dialog.destroy()

        ttk.Button(dialog, text="Agregar", command=confirm).pack(pady=(0, 10))

    def remove_range(self) -> None:
        selection = self.range_tree.selection()
        if not selection:
            messagebox.showinfo("Eliminar rango", "Seleccione un rango para eliminarlo.")
            return
        for item in selection:
            self.range_tree.delete(item)

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

        planes = list(self.plan_listbox.get(0, tk.END))
        rangos = []
        for item in self.range_tree.get_children():
            edad_min, edad_max, valor = self.range_tree.item(item, "values")
            rangos.append(
                AgeRange(
                    float(edad_min),
                    float(edad_max),
                    float(valor),
                )
            )

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
            rangos=rangos,
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

        self.plan_listbox.delete(0, tk.END)
        for plan in config.planes:
            self.plan_listbox.insert(tk.END, plan)

        for item in self.range_tree.get_children():
            self.range_tree.delete(item)
        for rango in config.rangos:
            self.range_tree.insert(
                "",
                tk.END,
                values=(
                    f"{rango.edad_min:.0f}" if rango.edad_min.is_integer() else f"{rango.edad_min}",
                    f"{rango.edad_max:.0f}" if rango.edad_max.is_integer() else f"{rango.edad_max}",
                    f"{rango.valor:.2f}",
                ),
            )

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

        # Validación de planes
        if config.planes and "Plan" in result.columns:
            result["Plan_Valido"] = result["Plan"].isin(config.planes)
        elif "Plan" in result.columns:
            result["Plan_Valido"] = True

        # Calcular valor por rango de edad
        if config.rangos and "Edad" in result.columns:
            rangos = sorted(config.rangos, key=lambda r: (r.edad_min, r.edad_max))

            def obtener_valor(edad: Any) -> Optional[float]:
                try:
                    edad_float = float(edad)
                except (TypeError, ValueError):
                    return None
                for rango in rangos:
                    if rango.edad_min <= edad_float <= rango.edad_max:
                        return rango.valor
                return None

            result["Valor_Rango"] = result["Edad"].apply(obtener_valor)
        else:
            result["Valor_Rango"] = None

        # Calcular tarifa final utilizando cobroFM y T_Congelada
        if config.cobro_fm.lower() == "si":
            result["Tarifa_Final"] = result["Valor_Rango"].fillna(0) * (1 + config.t_congelada)
        else:
            result["Tarifa_Final"] = result["Valor_Rango"]

        result.reset_index(drop=True, inplace=True)
        return result

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
