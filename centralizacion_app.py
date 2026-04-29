from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional

from centralizacion_remuneraciones import run_pipeline

# ── Constantes ──────────────────────────────────────────────────────────────
MESES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]
PROCESS_OPTIONS = [
    "Proceso Normal",
    "Proceso 1", "Proceso 2", "Proceso 3", "Proceso 4",
    "Proceso 5", "Proceso 6", "Proceso 7", "Proceso 8", "Proceso 9",
]
_DIC_DEFAULT = (
    Path(r"C:\Users\wilson.rojas\OneDrive - SERVICIO LOCAL DE EDUCACIÓN PÚBLICA DE PETORCA")
    / "SAF_SPYCG SLEP Petorca - Documentos"
    / "Diccionario Haberes por Fuente Financiamiento.xlsx"
)


# ── Widgets reutilizables ────────────────────────────────────────────────────

class SinglePath(ttk.Frame):
    """Fila de ruta única con label + entry + botón buscar."""

    def __init__(self, master, title: str, filetypes, directory: bool = False,
                 label_width: int = 26, entry_width: int = 80):
        super().__init__(master)
        self.filetypes = filetypes
        self.directory = directory
        ttk.Label(self, text=title, width=label_width, anchor="w").grid(
            row=0, column=0, sticky="w")
        self.var = tk.StringVar()
        ttk.Entry(self, textvariable=self.var, width=entry_width).grid(
            row=0, column=1, sticky="ew", padx=6)
        ttk.Button(self, text="Buscar", command=self.pick, width=8).grid(
            row=0, column=2, sticky="w")
        self.columnconfigure(1, weight=1)

    def pick(self):
        if self.directory:
            chosen = filedialog.askdirectory()
        else:
            chosen = filedialog.askopenfilename(filetypes=self.filetypes)
        if chosen:
            self.var.set(chosen)


class ProcessGroup(ttk.LabelFrame):
    """
    Bloque para un proceso del mes.
    Contiene: selector de nombre de proceso + 4 selectores de archivo.
    """

    def __init__(self, master, initial_name: str = "Proceso Normal",
                 on_remove=None, **kwargs):
        super().__init__(master, padding=8, **kwargs)
        self._on_remove = on_remove

        # ── Fila superior: nombre del proceso + botón quitar ──
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(0, 6))
        ttk.Label(top, text="Proceso:", font=("", 9, "bold")).pack(side="left")
        self.process_var = tk.StringVar(value=initial_name)
        cb = ttk.Combobox(
            top, textvariable=self.process_var,
            values=PROCESS_OPTIONS, width=22, state="normal",
        )
        cb.pack(side="left", padx=8)
        ttk.Label(top, text="(puedes escribir un nombre personalizado)",
                  foreground="gray").pack(side="left")
        if on_remove:
            ttk.Button(top, text="✕ Quitar proceso",
                       command=on_remove).pack(side="right")

        # ── Grilla de archivos ──
        grid = ttk.Frame(self)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)

        self._maestro = self._file_row(grid, "Maestro (*):", 0)
        self._gasto   = self._file_row(grid, "InformeGasto (*):", 1)
        self._asiento = self._file_row(grid, "Asiento (opcional):", 2)
        self._central = self._file_row(grid, "Central CAS (opcional):", 3)

    # ── helpers ──
    def _file_row(self, parent: ttk.Frame, label: str, row: int) -> tk.StringVar:
        ttk.Label(parent, text=label, width=26, anchor="w").grid(
            row=row, column=0, sticky="w", pady=2)
        var = tk.StringVar()
        ttk.Entry(parent, textvariable=var, width=78).grid(
            row=row, column=1, sticky="ew", padx=(4, 4))
        ttk.Button(parent, text="Seleccionar", width=12,
                   command=lambda v=var: self._pick(v)).grid(row=row, column=2)
        return var

    def _pick(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xls *.xlsx")])
        if path:
            var.set(path)

    # ── propiedades públicas ──
    @property
    def name(self) -> str:
        return self.process_var.get().strip() or "Proceso Normal"

    def is_valid(self) -> bool:
        return bool(self._maestro.get().strip() and self._gasto.get().strip())

    def get_files(self) -> Dict[str, str]:
        return {
            "name":    self.name,
            "maestro": self._maestro.get().strip(),
            "gasto":   self._gasto.get().strip(),
            "asiento": self._asiento.get().strip(),
            "central": self._central.get().strip(),
        }


class ScrollableFrame(ttk.Frame):
    """Frame con scrollbar vertical para contener widgets dinámicos."""

    def __init__(self, master, height: int = 380, **kwargs):
        super().__init__(master, **kwargs)
        self._canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0,
                                 height=height)
        sb = ttk.Scrollbar(self, orient="vertical",
                           command=self._canvas.yview)
        self.inner = ttk.Frame(self._canvas)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")
        self._canvas.configure(yscrollcommand=sb.set)
        self._canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_configure)
        self._canvas.bind("<Configure>", self._on_canvas_resize)
        self._canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_configure(self, _):
        self._canvas.configure(scrollregion=self._canvas.bbox("all"))

    def _on_canvas_resize(self, event):
        self._canvas.itemconfig(self._win_id, width=event.width)

    def _on_mousewheel(self, event):
        self._canvas.yview_scroll(-1 * (event.delta // 120), "units")


# ── Aplicación principal ─────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Centralización de Remuneraciones — SLEP Petorca")
        self.geometry("1200x940")
        self.minsize(1000, 700)
        self._groups: List[ProcessGroup] = []
        self._build_ui()

    # ── Construcción de la UI ──────────────────────────────────────────────

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        self._build_header(root)
        self._build_processes(root)
        self._build_aux(root)
        self._build_actions(root)
        self._set_defaults()

    def _build_header(self, parent):
        frm = ttk.LabelFrame(parent, text="Período y Área", padding=8)
        frm.pack(fill="x", pady=(0, 8))

        ttk.Label(frm, text="Área:").grid(row=0, column=0, sticky="w")
        self.area_var = tk.StringVar(value="Educacion")
        ttk.Combobox(frm, textvariable=self.area_var,
                     values=["Educacion", "Jardines"], width=16,
                     state="readonly").grid(row=0, column=1, padx=(4, 18))

        ttk.Label(frm, text="Mes:").grid(row=0, column=2, sticky="w")
        self.month_var = tk.StringVar(value="ABRIL")
        ttk.Combobox(frm, textvariable=self.month_var, values=MESES,
                     width=14, state="readonly").grid(row=0, column=3, padx=(4, 18))

        ttk.Label(frm, text="Año:").grid(row=0, column=4, sticky="w")
        self.year_var = tk.StringVar(value="2026")
        ttk.Combobox(frm, textvariable=self.year_var,
                     values=[str(y) for y in range(2025, 2036)],
                     width=8, state="readonly").grid(row=0, column=5, padx=4)

    def _build_processes(self, parent):
        outer = ttk.LabelFrame(parent, text="Procesos del mes", padding=8)
        outer.pack(fill="both", expand=True, pady=(0, 8))

        # Barra de botones "Agregar proceso"
        btn_bar = ttk.Frame(outer)
        btn_bar.pack(fill="x", pady=(0, 6))
        ttk.Label(btn_bar, text="Agregar:").pack(side="left")
        for name in ["Proceso Normal", "Proceso 1", "Proceso 2",
                     "Proceso 3", "Proceso 4"]:
            short = name.replace("Proceso ", "P.") if "Normal" not in name else "P.Normal"
            ttk.Button(btn_bar, text=f"+ {short}", width=11,
                       command=lambda n=name: self._add_group(n)
                       ).pack(side="left", padx=3)
        ttk.Label(btn_bar,
                  text="  ← Cada proceso necesita su Maestro e InformeGasto",
                  foreground="gray").pack(side="left", padx=8)

        # Área scrollable con los grupos
        self._scroll = ScrollableFrame(outer, height=360)
        self._scroll.pack(fill="both", expand=True)

        # Grupo inicial
        self._add_group("Proceso Normal")

    def _build_aux(self, parent):
        frm = ttk.LabelFrame(parent, text="Archivos auxiliares", padding=8)
        frm.pack(fill="x", pady=(0, 8))
        frm.columnconfigure(0, weight=1)

        self.mapeo_excel = SinglePath(frm, "Mapeo maestro cuentas:",
                                      [("Excel", "*.xlsx")])
        self.mapeo_excel.grid(row=0, column=0, sticky="ew", pady=3)

        self.diccionario = SinglePath(frm, "Diccionario haberes:",
                                      [("Excel", "*.xlsx")])
        self.diccionario.grid(row=1, column=0, sticky="ew", pady=3)

        self.salida_dir = SinglePath(frm, "Carpeta de salida:",
                                     [], directory=True)
        self.salida_dir.grid(row=2, column=0, sticky="ew", pady=3)

    def _build_actions(self, parent):
        acts = ttk.Frame(parent)
        acts.pack(fill="x", pady=(4, 4))
        self.generate_btn = ttk.Button(
            acts, text="▶  Generar centralización", width=26,
            command=self.generate)
        self.generate_btn.pack(side="left")
        ttk.Button(acts, text="Cerrar", command=self.destroy,
                   width=10).pack(side="right")

        self.status_var = tk.StringVar(value="Listo.")
        ttk.Label(parent, textvariable=self.status_var,
                  foreground="navy").pack(anchor="w", pady=(2, 2))

        self.log = tk.Text(parent, height=9, state="normal",
                           font=("Consolas", 9))
        self.log.pack(fill="both", expand=False)

    def _set_defaults(self):
        proj = Path(__file__).resolve().parent
        self.salida_dir.var.set(str(proj / "salidas"))
        self.mapeo_excel.var.set(str(proj / "mapeo_maestro.xlsx"))
        if _DIC_DEFAULT.exists():
            self.diccionario.var.set(str(_DIC_DEFAULT))

    # ── Gestión de grupos ──────────────────────────────────────────────────

    def _add_group(self, name: str = "Proceso Normal"):
        grp = ProcessGroup(
            self._scroll.inner,
            initial_name=name,
            on_remove=None,   # Se asigna abajo para capturar la ref correcta
        )
        # Asignar el callback de quitar DESPUÉS de crear para capturar grp
        btn_frame = grp.winfo_children()[0]   # el ttk.Frame "top"
        for child in btn_frame.winfo_children():
            if isinstance(child, ttk.Button) and "Quitar" in (child.cget("text") or ""):
                child.configure(command=lambda g=grp: self._remove_group(g))

        # Recrear con on_remove correcto
        grp.destroy()
        grp = ProcessGroup(
            self._scroll.inner,
            initial_name=name,
            on_remove=lambda g=None: self._remove_group(grp),
        )
        grp.pack(fill="x", pady=4, padx=2)
        self._groups.append(grp)
        # Scroll al final
        self._scroll._canvas.update_idletasks()
        self._scroll._canvas.yview_moveto(1.0)

    def _remove_group(self, grp: ProcessGroup):
        if len(self._groups) <= 1:
            messagebox.showwarning(
                "Aviso", "Debe haber al menos un proceso cargado.")
            return
        self._groups.remove(grp)
        grp.destroy()

    # ── Generación ────────────────────────────────────────────────────────

    def append_log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def generate(self):
        valid = [g for g in self._groups if g.is_valid()]
        if not valid:
            messagebox.showwarning(
                "Faltan archivos",
                "Cada proceso necesita al menos Maestro e InformeGasto.")
            return

        self.generate_btn.configure(state="disabled")
        self.status_var.set("Generando centralización…")
        self.append_log("─" * 60)
        self.append_log("Iniciando proceso…")
        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        try:
            output_dir = Path(self.salida_dir.var.get())
            output_dir.mkdir(parents=True, exist_ok=True)

            area  = self.area_var.get().strip()
            mes   = self.month_var.get().strip().upper()
            anio  = self.year_var.get().strip()
            period_label = f"{mes} {anio}"
            title_text = (
                f"CENTRALIZACION REMUNERACIONES {area.upper()} {period_label}"
            )
            output_file = (
                output_dir
                / f"centralizacion_{area.lower()}_{period_label.replace(' ', '_').lower()}.xlsx"
            )

            mapping_excel = Path(self.mapeo_excel.var.get().strip()
                                 or output_dir / "mapeo_maestro.xlsx")
            mapping_csv = mapping_excel.with_suffix(".csv")

            # Construir listas + mapa de proceso por archivo
            valid_groups = [g for g in self._groups if g.is_valid()]
            process_files: List[str] = []
            gasto_files:   List[str] = []
            asiento_files: List[str] = []
            central_files: List[str] = []
            pnm: Dict[str, str] = {}   # {ruta: nombre_proceso}

            for g in valid_groups:
                files = g.get_files()
                pname = files["name"]
                self.after(0, lambda n=pname: self.append_log(
                    f"  Cargando: {n}"))
                for key, lst in [("maestro",  process_files),
                                 ("gasto",    gasto_files),
                                 ("asiento",  asiento_files),
                                 ("central",  central_files)]:
                    p = files.get(key, "")
                    if p:
                        lst.append(p)
                        pnm[p] = pname

            salida, mapping_csv_path, mapping_xlsx_path, new_count, pending = run_pipeline(
                process_files=process_files,
                gasto_files=gasto_files,
                central_files=central_files,
                cuentas_pdf="",
                salida=str(output_file),
                mapeo=str(mapping_csv),
                mapeo_excel=str(mapping_excel),
                title_text=title_text,
                asiento_files=asiento_files,
                diccionario_path=self.diccionario.var.get().strip(),
                process_name_map=pnm,
            )

            self.after(0, lambda: self._done(
                salida=str(salida),
                mapping_xlsx=str(mapping_xlsx_path),
                mapping_csv=str(mapping_csv_path),
                new_count=new_count,
                pending=pending,
                n_procesos=len(valid_groups),
            ))

        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            self.after(0, lambda: self._failed(exc, tb))

    def _done(self, salida: str, mapping_xlsx: str, mapping_csv: str,
              new_count: int, pending: int, n_procesos: int):
        self.generate_btn.configure(state="normal")
        self.status_var.set("✔ Centralización completada.")
        self.append_log(f"Procesos consolidados: {n_procesos}")
        self.append_log(f"Archivo generado: {salida}")
        self.append_log(f"Mapeo maestro:    {mapping_xlsx}")
        if new_count:
            self.append_log(
                f"⚠  Códigos NUEVOS agregados al mapeo: {new_count}  — completa sus cuentas contables.")
        else:
            self.append_log("Mapeo maestro sin cambios (sin códigos nuevos).")
        self.append_log(f"Cuentas pendientes: {pending}")

        if new_count and pending:
            msg = (
                f"Centralización de {n_procesos} proceso(s) generada.\n\n"
                f"Se agregaron {new_count} código(s) nuevo(s) al mapeo maestro.\n"
                f"Quedan {pending} cuenta(s) por completar en:\n{mapping_xlsx}"
            )
        elif pending:
            msg = (
                f"Centralización lista ({n_procesos} proceso(s)).\n\n"
                f"Quedan {pending} cuenta(s) sin mapear en:\n{mapping_xlsx}"
            )
        else:
            msg = (
                f"Centralización lista ({n_procesos} proceso(s)).\n"
                f"Todas las cuentas están mapeadas.\n\n{salida}"
            )
        messagebox.showinfo("Completado", msg)

    def _failed(self, exc: Exception, tb: str = ""):
        self.generate_btn.configure(state="normal")
        self.status_var.set("✗ Error al generar.")
        self.append_log(f"ERROR: {exc}")
        if tb:
            self.append_log(tb)
        messagebox.showerror("Error", str(exc))


if __name__ == "__main__":
    App().mainloop()
