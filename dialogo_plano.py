#dialogo_plano.py
import json
import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

RUTA_CONFIG = os.path.join(os.getenv("APPDATA") or os.path.expanduser("~"), "ExtractorExcel", "config_plano.json")
NOMBRES_CLASE = {
    "EGRESO": "Egresos",
    "NOMINA": "Pagos de nómina",
    "RECIBO": "Recibos de caja",
    "BANCARIO": "Gastos/ingresos bancarios",
    "SEGURIDAD_SOCIAL": "Seguridad social",
}
PATRON_TIPO = re.compile(r"^[A-Z0-9]{1,5}$")
PATRON_CODPUC = re.compile(r"^\d{4,12}$")
PATRON_TERCERO = re.compile(r"^\d{5,15}(?:-\d)?$")
FUENTE = ("Segoe UI", 10)


def _cargar_config():
    try:
        with open(RUTA_CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def _guardar_config(cfg):
    try:
        os.makedirs(os.path.dirname(RUTA_CONFIG), exist_ok=True)
        with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def _moneda(valor):
    return f"{valor:,.2f}" if pd.notna(valor) else ""


def _fecha(ano, mes, dia):
    try:
        return f"{int(dia):02d}/{int(mes):02d}/{int(ano)}"
    except (TypeError, ValueError):
        return ""


class DialogoPlano(tk.Toplevel):
    COLUMNAS = (
        ("tipo", "Comprobante", 90, "center"),
        ("numero", "Número", 95, "center"),
        ("fecha", "Fecha", 85, "center"),
        ("detalle", "Detalle", 300, "w"),
        ("codpuc", "Cuenta contable", 110, "center"),
        ("tercero", "Tercero", 110, "center"),
        ("nombre", "Nombre", 200, "w"),
        ("debito", "Débito", 110, "e"),
        ("credito", "Crédito", 110, "e"),
        ("docrela", "Docrela", 90, "center"),
    )
    EDITABLES = {"tipo": "Comprobante", "codpuc": "Cuenta contable", "tercero": "Tercero", "docrela": "Docrela"}
    BUSQUEDA = ("detalle", "tipo", "codpuc", "tercero", "nombre", "docrela")

    def __init__(self, parent, generador, lin):
        super().__init__(parent)
        self.title("Parámetros del Archivo Plano")
        self.geometry("1300x680")
        self.minsize(1050, 540)
        self.transient(parent)
        self.grab_set()

        self.gen = generador
        self.lin = lin
        self._clave_consec = getattr(generador, "CLAVE_CONSECUTIVO", "consecutivo")
        self.predeterminados = dict(generador.COMPROBANTES_PREDETERMINADOS)
        self.tipo_clase = dict(self.predeterminados)
        self.cfg = _cargar_config()
        self.resultado = None
        self.nit_aplicado = None
        self._pendiente = None

        self.var_nit = tk.StringVar(value=getattr(generador, "nit_empresa", "") or self.cfg.get("ultimo_nit", ""))
        self.var_consec = tk.StringVar(value="1")
        self.vars_clase = {c: tk.StringVar(value=t) for c, t in self.tipo_clase.items()}
        self.var_buscar = tk.StringVar()
        self.var_filtro = tk.StringVar(value="Todas")
        self.var_cuenta = tk.StringVar(value="Todas")
        self.var_sin_tercero = tk.BooleanVar(value=False)
        self.var_campo = tk.StringVar(value=self.EDITABLES["tipo"])
        self.var_valor = tk.StringVar()
        self.var_sel = tk.StringVar(value="0 seleccionados")
        self.var_rango = tk.StringVar()
        self.var_totales = tk.StringVar()

        self._construir_ui()
        self._aplicar_nit()

        for var in (self.var_nit, self.var_consec, self.var_buscar):
            var.trace_add("write", lambda *_: self._programar())
        self.var_campo.trace_add("write", lambda *_: self._refrescar_opciones())
        self.bind("<Control-Return>", self._confirmar)
        self.bind("<Escape>", lambda e: self.destroy())
        self.ent_nit.focus_set()

    def _construir_ui(self):
        raiz = ttk.Frame(self, padding=12)
        raiz.pack(fill=tk.BOTH, expand=True)
        raiz.columnconfigure(1, weight=1)
        raiz.rowconfigure(1, weight=1)

        sup = ttk.Frame(raiz)
        sup.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 10))
        ttk.Label(sup, text="NIT de la Empresa (Tercero por defecto):", font=FUENTE).pack(side=tk.LEFT)
        self.ent_nit = ttk.Entry(sup, textvariable=self.var_nit, font=FUENTE, width=20)
        self.ent_nit.pack(side=tk.LEFT, padx=8)
        self.ent_nit.bind("<Return>", self._aplicar_nit)
        self.ent_nit.bind("<FocusOut>", self._aplicar_nit)
        ttk.Label(sup, text="Comprobantes y consecutivo se recuerdan por NIT.", foreground="gray").pack(side=tk.LEFT, padx=8)

        izq = ttk.Frame(raiz)
        izq.grid(row=1, column=0, sticky="ns", padx=(0, 12))

        frm_clases = ttk.LabelFrame(izq, text="Comprobante por clase", padding=10)
        frm_clases.pack(fill=tk.X)
        self.cmb_clase = {}
        self.lbl_conteo = {}
        for r, clase in enumerate(self.predeterminados):
            ttk.Label(frm_clases, text=NOMBRES_CLASE.get(clase, clase)).grid(row=r, column=0, sticky="w", pady=2)
            cmb = ttk.Combobox(frm_clases, textvariable=self.vars_clase[clase], width=8)
            cmb.grid(row=r, column=1, padx=6, pady=2)
            for evento in ("<<ComboboxSelected>>", "<Return>", "<FocusOut>"):
                cmb.bind(evento, lambda e, c=clase: self._cambiar_clase(c) and self._refrescar())
            self.lbl_conteo[clase] = ttk.Label(frm_clases, foreground="gray")
            self.lbl_conteo[clase].grid(row=r, column=2, sticky="e")
            self.cmb_clase[clase] = cmb
        if self.gen.CLASE_BANCARIO in self.predeterminados:
            ttk.Button(frm_clases, text="Pasar retiros a CC", command=self._pasar_retiros).grid(
                row=len(self.predeterminados), column=0, columnspan=3, sticky="ew", pady=(8, 0))

        frm_consec = ttk.LabelFrame(izq, text="Consecutivo", padding=10)
        frm_consec.pack(fill=tk.X, pady=(10, 0))
        ttk.Label(frm_consec, text="Primer consecutivo:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm_consec, textvariable=self.var_consec, width=12, justify="right").grid(row=0, column=1, padx=6)
        ttk.Label(frm_consec, textvariable=self.var_rango, foreground="gray").grid(row=1, column=0, columnspan=2, sticky="w", pady=(4, 0))

        ttk.Button(izq, text="Generar archivo plano  (Ctrl+Enter)", command=self._confirmar).pack(side=tk.BOTTOM, fill=tk.X)

        der = ttk.Frame(raiz)
        der.grid(row=1, column=1, sticky="nsew")
        der.columnconfigure(0, weight=1)
        der.rowconfigure(2, weight=1)

        filtros = ttk.Frame(der)
        filtros.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 4))
        ttk.Label(filtros, text="Buscar:").pack(side=tk.LEFT)
        ttk.Entry(filtros, textvariable=self.var_buscar, width=30).pack(side=tk.LEFT, padx=6)
        ttk.Label(filtros, text="Clase:").pack(side=tk.LEFT, padx=(10, 0))
        cmb_filtro = ttk.Combobox(
            filtros, textvariable=self.var_filtro, state="readonly", width=24,
            values=["Todas"] + [NOMBRES_CLASE.get(c, c) for c in self.predeterminados],
        )
        cmb_filtro.pack(side=tk.LEFT, padx=6)
        cmb_filtro.bind("<<ComboboxSelected>>", lambda e: self._refrescar())
        ttk.Label(filtros, text="Cuenta:").pack(side=tk.LEFT, padx=(10, 0))
        self.cmb_cuenta = ttk.Combobox(filtros, textvariable=self.var_cuenta, state="readonly", width=12)
        self.cmb_cuenta.pack(side=tk.LEFT, padx=6)
        self.cmb_cuenta.bind("<<ComboboxSelected>>", lambda e: self._refrescar())
        ttk.Checkbutton(filtros, text="Terceros vacíos", variable=self.var_sin_tercero,
                        command=self._refrescar).pack(side=tk.LEFT, padx=(10, 0))
        ttk.Button(filtros, text="Limpiar filtros", command=self._limpiar_filtros).pack(side=tk.RIGHT)
        ttk.Label(der, text="Doble clic en Comprobante, Cuenta, Tercero o Docrela para editar · vacío = automático",
                  foreground="gray").grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 4))

        self.tree = ttk.Treeview(der, columns=[c[0] for c in self.COLUMNAS], show="headings", selectmode="extended")
        for col, titulo, ancho, ancla in self.COLUMNAS:
            self.tree.heading(col, text=titulo)
            self.tree.column(col, width=ancho, minwidth=60, anchor=ancla, stretch=(col == "detalle"))
        self.tree.tag_configure("par", background="#f2f5f9")
        self.tree.tag_configure("manual", foreground="#1a5fb4")
        self.tree.grid(row=2, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(der, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_y.grid(row=2, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(der, orient=tk.HORIZONTAL, command=self.tree.xview)
        scroll_x.grid(row=3, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.bind("<<TreeviewSelect>>", lambda e: self.var_sel.set(f"{len(self.tree.selection())} seleccionados"))
        self.tree.bind("<Control-a>", self._seleccionar_todo)
        self.tree.bind("<Double-1>", self._editar_celda)

        acciones = ttk.Frame(der)
        acciones.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(acciones, text="Seleccionar todo", command=self._seleccionar_todo).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(acciones, textvariable=self.var_sel, foreground="gray").pack(side=tk.LEFT)
        self.lbl_totales = ttk.Label(acciones, textvariable=self.var_totales, foreground="gray")
        self.lbl_totales.pack(side=tk.LEFT, padx=12)
        ttk.Button(acciones, text="Restablecer", command=self._restablecer).pack(side=tk.RIGHT)
        ttk.Button(acciones, text="Aplicar a selección", command=self._aplicar_seleccion).pack(side=tk.RIGHT, padx=6)
        self.cmb_valor = ttk.Combobox(acciones, textvariable=self.var_valor, width=14)
        self.cmb_valor.pack(side=tk.RIGHT)
        self.cmb_valor.bind("<Return>", self._aplicar_seleccion)
        ttk.Combobox(acciones, textvariable=self.var_campo, state="readonly", width=15,
                     values=list(self.EDITABLES.values())).pack(side=tk.RIGHT, padx=6)
        ttk.Label(acciones, text="Asignar:").pack(side=tk.RIGHT)

    @staticmethod
    def _normalizar(tipo):
        tipo = str(tipo or "").strip().upper()
        return tipo if PATRON_TIPO.match(tipo) else None

    def _avisar(self, titulo, mensaje):
        messagebox.showwarning(titulo, mensaje, parent=self)
        return False

    def _primer(self):
        valor = self.var_consec.get().strip()
        return int(valor) if valor.isdigit() and int(valor) > 0 else None

    def _programar(self):
        if self._pendiente:
            self.after_cancel(self._pendiente)
        self._pendiente = self.after(150, self._refrescar)

    def _calcular(self, primer=None):
        calc = self.gen.calcular(self.lin, self.var_nit.get().strip(), primer or self._primer() or 1)
        return calc if "nombre" in calc.columns else calc.assign(nombre="")

    def _aplicar_nit(self, _=None):
        nit = self.var_nit.get().strip()
        if nit == self.nit_aplicado:
            return
        self.nit_aplicado = nit
        empresa = self.cfg.get("empresas", {}).get(nit, {})
        clases = empresa.get("clases", {})
        for clase, defecto in self.predeterminados.items():
            self._cambiar_clase(clase, clases.get(clase, defecto))
        self.var_consec.set(str(empresa.get(self._clave_consec, 1)))
        self._refrescar()

    def _cambiar_clase(self, clase, nuevo=None):
        viejo = self.tipo_clase[clase]
        tipo = self._normalizar(self.vars_clase[clase].get() if nuevo is None else nuevo)
        self.vars_clase[clase].set(tipo or viejo)
        if not tipo:
            return self._avisar("Comprobante inválido", "El comprobante debe tener de 1 a 5 caracteres alfanuméricos (ej. EG1).")
        if tipo != viejo:
            self.tipo_clase[clase] = tipo
            mask = (self.lin["clase"] == clase) & (self.lin["tipo"] == viejo)
            self.lin.loc[mask, "tipo"] = tipo
        return True

    def _refrescar(self, ver=None):
        self._pendiente = None
        primer = self._primer()
        calc = self._calcular(primer)
        n = calc["mov"].nunique()
        self.var_rango.set(f"{n} comprobantes → hasta {primer + n - 1}" if primer else f"{n} comprobantes · consecutivo inválido")

        par = pd.Series(pd.factorize(calc["mov"])[0] % 2 == 1, index=calc.index)
        manual = (
            (calc["codpuc_manual"] != "") | (calc["tercero_manual"] != "")
            | (calc["tipo"] != calc["clase"].map(self.tipo_clase))
        )

        vista = calc
        clase = next((c for c in self.predeterminados if NOMBRES_CLASE.get(c, c) == self.var_filtro.get()), None)
        if clase:
            vista = vista[vista["clase"] == clase]
        if self.var_cuenta.get() != "Todas":
            vista = vista[vista["codpuc"] == self.var_cuenta.get()]
        if self.var_sin_tercero.get():
            vista = vista[vista["tercero"].str.strip() == ""]
        texto = self.var_buscar.get().strip()
        if texto:
            coincide = pd.Series(False, index=vista.index)
            for col in self.BUSQUEDA:
                coincide |= vista[col].astype(str).str.contains(texto, case=False, regex=False, na=False)
            vista = vista[coincide]

        seleccion = set(self.tree.selection())
        posicion = self.tree.yview()[0]
        self.tree.delete(*self.tree.get_children())
        for i, tipo, num, ano, mes, dia, det, cod, ter, nom, deb, cre, doc in zip(
            vista.index, vista["tipo"], vista["numero"], vista["ano"], vista["mes"], vista["dia"], vista["detalle"],
            vista["codpuc"], vista["tercero"], vista["nombre"], vista["debito"], vista["credito"], vista["docrela"],
        ):
            tags = (("par",) if par[i] else ()) + (("manual",) if manual[i] else ())
            self.tree.insert("", tk.END, iid=str(i), tags=tags, values=(
                tipo, num, _fecha(ano, mes, dia), det, cod, ter, nom, _moneda(deb), _moneda(cre), doc,
            ))

        self.tree.selection_set([s for s in seleccion if self.tree.exists(s)])
        self.tree.yview_moveto(posicion)
        if ver and self.tree.exists(ver):
            self.tree.see(ver)
        self.var_sel.set(f"{len(self.tree.selection())} seleccionados")

        debitos, creditos = vista["debito"].sum(), vista["credito"].sum()
        self.var_totales.set(f"Débitos {debitos:,.2f} · Créditos {creditos:,.2f}")
        self.lbl_totales.config(foreground="#c01c28" if abs(debitos - creditos) > 0.005 else "gray")
        self.cmb_cuenta.config(values=["Todas"] + sorted(calc["codpuc"].unique()))
        conteo = calc.drop_duplicates("mov")["clase"].value_counts()
        for c, lbl in self.lbl_conteo.items():
            lbl.config(text=f"{int(conteo.get(c, 0))} mov.")
        self._refrescar_opciones()

    def _limpiar_filtros(self):
        self.var_filtro.set("Todas")
        self.var_cuenta.set("Todas")
        self.var_sin_tercero.set(False)
        self.var_buscar.set("")
        self._refrescar()

    def _pasar_retiros(self):
        bancario = self.gen.CLASE_BANCARIO
        mask = (self.lin["clase"] != bancario) & self.lin["detalle"].str.contains(
            self.gen.PATRON_RETIROS, case=False, regex=True, na=False)
        if not mask.any():
            messagebox.showinfo("Retiros", "No hay retiros pendientes por pasar a CC.", parent=self)
            return
        n = self.lin.loc[mask, "mov"].nunique()
        self.lin.loc[mask, "clase"] = bancario
        self.lin.loc[mask, "tipo"] = self.tipo_clase[bancario]
        self.lin.loc[mask & (self.lin["codpuc_contra"].fillna("") == ""), "codpuc_contra"] = self.gen.CODPUC_RETIROS
        self._refrescar()
        messagebox.showinfo("Retiros", f"{n} movimientos pasados a {self.tipo_clase[bancario]}.", parent=self)

    def _refrescar_opciones(self):
        guardados = {t for e in self.cfg.get("empresas", {}).values() for t in e.get("clases", {}).values()}
        opciones = sorted(set(self.predeterminados.values()) | set(self.lin["tipo"]) | guardados)
        for cmb in self.cmb_clase.values():
            cmb.config(values=opciones)
        self.cmb_valor.config(values=opciones if self.var_campo.get() == self.EDITABLES["tipo"] else [])

    def _seleccionar_todo(self, _=None):
        self.tree.selection_set(self.tree.get_children())
        return "break"

    def _editar_celda(self, event):
        iid = self.tree.identify_row(event.y)
        columna = self.tree.identify_column(event.x)
        if not iid or not columna:
            return
        campo = self.COLUMNAS[int(columna[1:]) - 1][0]
        if campo not in self.EDITABLES:
            return
        seleccion = self.tree.selection()
        filas = [int(i) for i in seleccion] if iid in seleccion else [int(iid)]
        x, y, w, h = self.tree.bbox(iid, campo)
        editor = ttk.Entry(self.tree, justify="center")
        editor.place(x=x, y=y, width=w, height=h)
        editor.insert(0, self.tree.set(iid, campo))
        editor.select_range(0, tk.END)
        editor.focus_set()

        def cerrar(guardar):
            if editor.winfo_exists():
                valor = editor.get()
                editor.destroy()
                if guardar:
                    self._asignar(campo, filas, valor)
            return "break"

        editor.bind("<Return>", lambda e: cerrar(True))
        editor.bind("<FocusOut>", lambda e: cerrar(True))
        editor.bind("<Escape>", lambda e: cerrar(False))

    def _asignar(self, campo, filas, valor):
        valor = str(valor or "").strip()
        if campo == "tipo":
            tipo = self._normalizar(valor)
            if not tipo:
                return self._avisar("Comprobante inválido", "El comprobante debe tener de 1 a 5 caracteres alfanuméricos (ej. EG1).")
            movs = self.lin.loc[filas, "mov"].unique()
            self.lin.loc[self.lin["mov"].isin(movs), "tipo"] = tipo
        elif campo == "codpuc":
            if valor and not PATRON_CODPUC.match(valor):
                return self._avisar("Cuenta inválida", "La cuenta contable debe tener de 4 a 12 dígitos.")
            self.lin.loc[filas, "codpuc_manual"] = valor
        elif campo == "tercero":
            if valor and not PATRON_TERCERO.match(valor):
                return self._avisar("Tercero inválido", "El tercero debe ser un NIT o cédula numérico (ej. 900123456 o 900123456-7).")
            self.lin.loc[filas, "tercero_manual"] = valor
        else:
            self.lin.loc[filas, "docrela"] = valor
        self._refrescar(ver=str(filas[0]))
        return True

    def _filas_seleccionadas(self):
        filas = [int(i) for i in self.tree.selection()]
        if not filas:
            messagebox.showinfo("Sin selección", "Seleccione una o más filas (Ctrl+A selecciona las visibles).", parent=self)
        return filas

    def _aplicar_seleccion(self, _=None):
        filas = self._filas_seleccionadas()
        if filas:
            campo = next(k for k, v in self.EDITABLES.items() if v == self.var_campo.get())
            self._asignar(campo, filas, self.var_valor.get())

    def _restablecer(self):
        filas = self._filas_seleccionadas()
        if not filas:
            return
        mask = self.lin["mov"].isin(self.lin.loc[filas, "mov"].unique())
        self.lin.loc[mask, "tipo"] = self.lin.loc[mask, "clase"].map(self.tipo_clase)
        self.lin.loc[filas, ["codpuc_manual", "tercero_manual", "docrela"]] = ""
        self._refrescar(ver=str(filas[0]))

    def _confirmar(self, _=None):
        if not all(self._cambiar_clase(c) for c in self.predeterminados):
            return
        primer = self._primer()
        if not primer:
            self._avisar("Consecutivo inválido", "Ingrese un primer consecutivo válido.")
            return
        if self._calcular(primer)["codpuc"].eq("").any():
            self._avisar("Cuenta faltante", "Existen líneas sin cuenta contable. Asígnela con doble clic en 'Cuenta contable'.")
            return
        self.resultado = {
            "lin": self.lin,
            "consecutivo": primer,
            "nit": self.var_nit.get().strip(),
            "clases": dict(self.tipo_clase),
            "clave": self._clave_consec,
        }
        self.destroy()


def _guardar_parametros(resultado):
    cfg = _cargar_config()
    empresa = cfg.setdefault("empresas", {}).setdefault(resultado["nit"], {})
    empresa.pop("consecutivos", None)
    empresa.setdefault("clases", {}).update(resultado["clases"])
    empresa[resultado.get("clave", "consecutivo")] = resultado["consecutivo"] + int(resultado["lin"]["mov"].nunique())
    cfg["ultimo_nit"] = resultado["nit"]
    _guardar_config(cfg)


def _con_carga(parent, texto, funcion):
    res = {}

    def trabajo():
        try:
            res["ok"] = funcion()
        except Exception as e:
            res["error"] = e

    ven = tk.Toplevel(parent)
    ven.title("Archivo Plano")
    ven.transient(parent)
    ven.resizable(False, False)
    ven.protocol("WM_DELETE_WINDOW", lambda: None)
    ttk.Label(ven, text=texto, font=FUENTE, padding=(20, 15, 20, 8)).pack()
    barra = ttk.Progressbar(ven, mode="indeterminate", length=300)
    barra.pack(padx=20, pady=(0, 18))
    barra.start(10)
    ven.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - ven.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - ven.winfo_height()) // 2
    ven.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    ven.grab_set()

    hilo = threading.Thread(target=trabajo, daemon=True)
    hilo.start()

    def revisar():
        if hilo.is_alive():
            ven.after(100, revisar)
        else:
            barra.stop()
            ven.destroy()

    ven.after(100, revisar)
    parent.wait_window(ven)
    if "error" in res:
        raise res["error"]
    return res.get("ok")


def generar_plano_con_dialogo(parent, excel_path, generador=None, confirmar=True, pedir_ruta=False):
    if confirmar and not messagebox.askyesno("Archivo Plano", "¿Desea ejecutar el archivo plano?", parent=parent):
        return None

    if generador is None:
        from archivo_plano import GeneradorArchivoPlano
        generador = GeneradorArchivoPlano(excel_path)

    try:
        lin = _con_carga(parent, "Leyendo documento y preparando movimientos...", generador.preparar)
    except Exception as e:
        messagebox.showerror("Error en Plano", f"No fue posible leer los movimientos:\n{e}", parent=parent)
        return None

    if lin is None or lin.empty:
        messagebox.showinfo("Archivo Plano", "El archivo no contiene movimientos con valor.", parent=parent)
        return None

    if getattr(generador, "advertencias", None):
        messagebox.showwarning("Archivo Plano", "\n\n".join(generador.advertencias), parent=parent)

    dialogo = DialogoPlano(parent, generador, lin)
    parent.wait_window(dialogo)
    resultado = dialogo.resultado
    if not resultado:
        return None

    ruta_destino = None
    if pedir_ruta:
        base = os.path.splitext(os.path.basename(excel_path))[0]
        ruta_destino = filedialog.asksaveasfilename(
            parent=parent,
            title="Guardar archivo plano como...",
            initialdir=os.path.dirname(excel_path),
            initialfile=f"{base}_PLANO.xls",
            defaultextension=".xls",
            filetypes=[("Excel 5.0/95", "*.xls")],
        )
        if not ruta_destino:
            return None
        ruta_destino = os.path.normpath(ruta_destino)

    try:
        ruta_plano = _con_carga(parent, "Generando archivo plano...", lambda: generador.generar(
            lin=resultado["lin"],
            nit_empresa=resultado["nit"],
            primer_consecutivo=resultado["consecutivo"],
            ruta_destino_final=ruta_destino,
        ))
    except Exception as e:
        messagebox.showerror("Error en Plano", f"Falló la generación del archivo plano:\n{e}", parent=parent)
        return None

    if not ruta_plano or not os.path.exists(ruta_plano):
        messagebox.showerror("Error en Plano", f"No se encontró el archivo plano generado:\n{ruta_plano}", parent=parent)
        return None

    _guardar_parametros(resultado)
    messagebox.showinfo("Archivo Plano", f"El archivo plano se ha generado en:\n{ruta_plano}", parent=parent)
    return ruta_plano