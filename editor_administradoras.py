#editor_administradoras.py
import json
import os
import re
import shutil
import tempfile
import tkinter as tk
from tkinter import messagebox, ttk

from dialogo_plano import _modal
from script_seguridad_social import RUTA_ADMINISTRADORAS, _normalizar

FUENTE = ("Segoe UI", 10)
VALIDACIONES = {
    "NIT": (re.compile(r"^\d{5,15}$"), "El NIT debe tener de 5 a 15 dígitos, sin dígito de verificación."),
    "CUENTA": (re.compile(r"^\d{4,12}$"), "La cuenta contable debe tener de 4 a 12 dígitos."),
    "CODIGO": (re.compile(r"^[A-Z0-9-]{2,10}$"), "El código PILA admite letras, números y guion (ej. EPS010, 25-14)."),
}
OBLIGATORIAS = ("NOMBRE", "NIT")
TIPOS = ("NOMBRE", "NIT", "CUENTA", "CODIGO", "CATEGOR")


def _tipo(columna):
    norm = _normalizar(columna)
    return next((t for t in TIPOS if t in norm), None)


def _limpiar(tipo, valor):
    valor = str(valor or "").strip()
    if tipo in ("NIT", "CUENTA"):
        return re.sub(r"[.\s]", "", valor)
    if tipo == "CODIGO":
        return re.sub(r"\s", "", valor).upper()
    return valor


def _leer(ruta):
    with open(ruta, encoding="utf-8-sig") as f:
        raw = json.load(f)
    clave = None
    if isinstance(raw, dict):
        clave = next((k for k, v in raw.items() if isinstance(v, list) and (not v or isinstance(v[0], dict))), None)
        if clave is None:
            raise ValueError("administradoras.json no contiene una lista de administradoras.")
        registros = raw[clave]
    elif isinstance(raw, list):
        registros = raw
    else:
        raise ValueError("Formato de administradoras.json no reconocido.")
    return raw, clave, [dict(r) for r in registros if isinstance(r, dict)], os.path.getmtime(ruta)


class EditorAdministradoras(tk.Toplevel):
    def __init__(self, parent, ruta=RUTA_ADMINISTRADORAS):
        self.ruta = ruta
        datos = _leer(ruta)
        super().__init__(parent)
        self.title("Administradoras de seguridad social")
        self.geometry("1050x600")
        self.minsize(800, 420)
        self.protocol("WM_DELETE_WINDOW", self._cerrar)

        self.var_buscar = tk.StringVar()
        self.var_estado = tk.StringVar()
        self._orden = (None, False)
        self._construir_ui()
        self._aplicar(*datos)
        self.var_buscar.trace_add("write", lambda *_: self._refrescar())
        self.bind("<Control-s>", lambda e: self._guardar())

    def _aplicar(self, raw, clave, registros, mtime):
        self.raw, self.clave, self.registros, self.mtime = raw, clave, registros, mtime
        columnas = []
        for r in registros:
            columnas += [c for c in r if c not in columnas]
        if not any(_tipo(c) == "CODIGO" for c in columnas):
            columnas.append("CODIGO")
        self.columnas = columnas
        self.cambios = False
        self.tree.config(columns=[f"c{i}" for i in range(len(columnas))])
        for i, col in enumerate(columnas):
            self.tree.heading(f"c{i}", text=col, command=lambda i=i: self._ordenar(i))
            ancho = 280 if _tipo(col) == "NOMBRE" else 120
            self.tree.column(f"c{i}", width=ancho, minwidth=60, anchor="w" if ancho > 120 else "center",
                             stretch=_tipo(col) == "NOMBRE")
        self._refrescar()

    def _construir_ui(self):
        raiz = ttk.Frame(self, padding=12)
        raiz.pack(fill=tk.BOTH, expand=True)
        raiz.columnconfigure(0, weight=1)
        raiz.rowconfigure(2, weight=1)

        sup = ttk.Frame(raiz)
        sup.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(sup, text="Buscar:", font=FUENTE).pack(side=tk.LEFT)
        ttk.Entry(sup, textvariable=self.var_buscar, width=35, font=FUENTE).pack(side=tk.LEFT, padx=6)
        ttk.Label(sup, textvariable=self.var_estado, foreground="gray").pack(side=tk.LEFT, padx=10)

        ttk.Label(raiz, text=f"Archivo: {self.ruta}", foreground="gray").grid(row=1, column=0, columnspan=2, sticky="w")

        self.tree = ttk.Treeview(raiz, show="headings", selectmode="extended")
        self.tree.grid(row=2, column=0, sticky="nsew", pady=(4, 0))
        self.tree.tag_configure("par", background="#f2f5f9")
        self.tree.tag_configure("incompleta", foreground="#c01c28")
        scroll_y = ttk.Scrollbar(raiz, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_y.grid(row=2, column=1, sticky="ns", pady=(4, 0))
        scroll_x = ttk.Scrollbar(raiz, orient=tk.HORIZONTAL, command=self.tree.xview)
        scroll_x.grid(row=3, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree.bind("<Double-1>", lambda e: self._editar())
        self.tree.bind("<Delete>", lambda e: self._eliminar())

        acciones = ttk.Frame(raiz)
        acciones.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(acciones, text="➕ Agregar", command=self._agregar).pack(side=tk.LEFT)
        ttk.Button(acciones, text="✏️ Editar", command=self._editar).pack(side=tk.LEFT, padx=6)
        ttk.Button(acciones, text="🗑 Eliminar", command=self._eliminar).pack(side=tk.LEFT)
        ttk.Label(acciones, text="En rojo: sin NIT, cuenta o código PILA.", foreground="#c01c28").pack(side=tk.LEFT, padx=12)
        ttk.Button(acciones, text="💾 Guardar  (Ctrl+S)", command=self._guardar).pack(side=tk.RIGHT)
        ttk.Button(acciones, text="⟳ Recargar", command=self._recargar).pack(side=tk.RIGHT, padx=6)

    def _valor(self, registro, col):
        v = registro.get(col, "")
        return "" if v is None else str(v)

    def _incompleta(self, registro):
        return any(not self._valor(registro, c).strip() for c in self.columnas if _tipo(c) in ("NIT", "CUENTA", "CODIGO"))

    def _ordenar(self, i):
        col, desc = self._orden
        self._orden = (i, not desc if col == i else False)
        self._refrescar()

    def _refrescar(self, ver=None):
        texto = _normalizar(self.var_buscar.get())
        indices = list(range(len(self.registros)))
        if texto:
            indices = [i for i in indices
                       if texto in _normalizar(" ".join(self._valor(self.registros[i], c) for c in self.columnas))]
        col, desc = self._orden
        if col is not None:
            indices.sort(key=lambda i: self._valor(self.registros[i], self.columnas[col]).upper(), reverse=desc)

        self.tree.delete(*self.tree.get_children())
        for n, i in enumerate(indices):
            r = self.registros[i]
            tags = (("par",) if n % 2 else ()) + (("incompleta",) if self._incompleta(r) else ())
            self.tree.insert("", tk.END, iid=str(i), tags=tags, values=[self._valor(r, c) for c in self.columnas])
        if ver is not None and self.tree.exists(str(ver)):
            self.tree.selection_set(str(ver))
            self.tree.see(str(ver))
        pendientes = sum(self._incompleta(r) for r in self.registros)
        self.var_estado.set(
            f"{len(indices)} de {len(self.registros)} administradoras · {pendientes} incompletas"
            + (" · ● cambios sin guardar" if self.cambios else "")
        )

    def _seleccion(self):
        sel = [int(i) for i in self.tree.selection()]
        if not sel:
            messagebox.showinfo("Administradoras", "Seleccione una administradora.", parent=self)
        return sel

    def _agregar(self):
        self._formulario(None)

    def _editar(self):
        sel = self._seleccion()
        if sel:
            self._formulario(sel[0])

    def _eliminar(self):
        sel = self._seleccion()
        if not sel:
            return
        col = next((c for c in self.columnas if _tipo(c) == "NOMBRE"), self.columnas[0])
        nombres = "\n- ".join(self._valor(self.registros[i], col) for i in sel[:10])
        if not messagebox.askyesno("Eliminar", f"¿Eliminar {len(sel)} administradora(s)?\n- {nombres}", parent=self):
            return
        for i in sorted(sel, reverse=True):
            del self.registros[i]
        self.cambios = True
        self._refrescar()

    def _formulario(self, indice):
        original = self.registros[indice] if indice is not None else {}
        ven = tk.Toplevel(self)
        ven.title("Editar administradora" if indice is not None else "Nueva administradora")
        ven.resizable(False, False)
        frm = ttk.Frame(ven, padding=14)
        frm.pack(fill=tk.BOTH, expand=True)

        variables = {}
        for r, col in enumerate(self.columnas):
            obligatoria = _tipo(col) in OBLIGATORIAS
            ttk.Label(frm, text=f"{col}{' *' if obligatoria else ''}", font=FUENTE).grid(row=r, column=0, sticky="w", pady=3)
            var = tk.StringVar(value=self._valor(original, col))
            ent = ttk.Entry(frm, textvariable=var, width=45, font=FUENTE)
            ent.grid(row=r, column=1, padx=(10, 0), pady=3)
            if r == 0:
                ent.focus_set()
            variables[col] = var
        ttk.Label(frm, text="Código PILA: el que aparece en la columna 'Código' de la planilla (ej. EPS010, 230301, 25-14).",
                  foreground="gray", wraplength=420).grid(row=len(self.columnas), column=0, columnspan=2, sticky="w", pady=(8, 0))

        def aceptar(_=None):
            valores = {c: _limpiar(_tipo(c), v.get()) for c, v in variables.items()}
            for col, valor in valores.items():
                tipo = _tipo(col)
                if tipo in OBLIGATORIAS and not valor:
                    return messagebox.showwarning("Dato faltante", f"El campo '{col}' es obligatorio.", parent=ven)
                regla = VALIDACIONES.get(tipo)
                if valor and regla and not regla[0].match(valor):
                    return messagebox.showwarning("Dato inválido", f"{col}: {regla[1]}", parent=ven)
            col_cod = next((c for c in self.columnas if _tipo(c) == "CODIGO"), None)
            col_nit = next((c for c in self.columnas if _tipo(c) == "NIT"), None)
            codigo = valores.get(col_cod, "")
            if codigo and col_nit:
                choque = [self._valor(r, col_nit) for i, r in enumerate(self.registros)
                          if i != indice and _limpiar("CODIGO", r.get(col_cod)) == codigo
                          and _limpiar("NIT", r.get(col_nit)) != valores[col_nit]]
                if choque:
                    return messagebox.showwarning(
                        "Código duplicado",
                        f"El código {codigo} ya está asignado al NIT {choque[0]}.\nUn código PILA corresponde a una sola administradora.",
                        parent=ven)
            col_cat = next((c for c in self.columnas if _tipo(c) == "CATEGOR"), None)
            if col_nit:
                nit = valores[col_nit]
                cat = _normalizar(valores.get(col_cat, ""))
                existente = next((r for i, r in enumerate(self.registros)
                                  if i != indice and _limpiar("NIT", r.get(col_nit)) == nit
                                  and _normalizar(r.get(col_cat) if col_cat else "") == cat), None)
                if existente:
                    col_nom = next((c for c in self.columnas if _tipo(c) == "NOMBRE"), col_nit)
                    return messagebox.showwarning(
                        "NIT duplicado",
                        f"El NIT {nit} ya existe para {self._valor(existente, col_nom)}"
                        + (f" ({cat})" if cat else "") + ".\nNo se puede crear otra administradora con el mismo NIT.",
                        parent=ven)
                
            registro = dict(original)
            for col, valor in valores.items():
                previo = original.get(col)
                if not valor and col not in original:
                    continue
                registro[col] = int(valor) if isinstance(previo, int) and not isinstance(previo, bool) and valor.isdigit() else valor
            if indice is None:
                self.registros.append(registro)
                nuevo = len(self.registros) - 1
            else:
                self.registros[indice] = registro
                nuevo = indice
            self.cambios = True
            ven.destroy()
            self._refrescar(ver=nuevo)

        botones = ttk.Frame(frm)
        botones.grid(row=len(self.columnas) + 1, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(botones, text="Cancelar", command=ven.destroy).pack(side=tk.RIGHT)
        ttk.Button(botones, text="Aceptar", command=aceptar).pack(side=tk.RIGHT, padx=6)
        ven.bind("<Return>", aceptar)
        ven.bind("<Escape>", lambda e: ven.destroy())
        _modal(ven, self)

    def _recargar(self):
        if self.cambios and not messagebox.askyesno("Recargar", "Se perderán los cambios sin guardar. ¿Continuar?", parent=self):
            return
        try:
            self._aplicar(*_leer(self.ruta))
        except Exception as e:
            messagebox.showerror("Administradoras", f"No se pudo leer el archivo:\n{e}", parent=self)

    def _guardar(self):
        if not self.cambios:
            messagebox.showinfo("Administradoras", "No hay cambios por guardar.", parent=self)
            return False
        try:
            if os.path.getmtime(self.ruta) != self.mtime and not messagebox.askyesno(
                "Archivo modificado",
                "Otro usuario modificó administradoras.json después de que usted lo abrió.\n\n"
                "Sí: sobrescribir con sus cambios.\nNo: cancelar (use Recargar para ver la versión actual).",
                icon="warning", parent=self,
            ):
                return False

            contenido = self.registros
            if self.clave is not None:
                self.raw[self.clave] = self.registros
                contenido = self.raw
            carpeta = os.path.dirname(self.ruta)
            shutil.copy2(self.ruta, self.ruta + ".bak")
            fd, temporal = tempfile.mkstemp(suffix=".tmp", dir=carpeta)
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as f:
                    json.dump(contenido, f, ensure_ascii=False, indent=2)
                os.replace(temporal, self.ruta)
            except Exception:
                if os.path.exists(temporal):
                    os.remove(temporal)
                raise
            self.mtime = os.path.getmtime(self.ruta)
        except Exception as e:
            messagebox.showerror("Administradoras", f"No se pudo guardar:\n{e}", parent=self)
            return False
        self.cambios = False
        self._refrescar()
        messagebox.showinfo("Administradoras", "Cambios guardados. Copia anterior: administradoras.json.bak", parent=self)
        return True

    def _cerrar(self):
        if self.cambios:
            respuesta = messagebox.askyesnocancel("Administradoras", "Hay cambios sin guardar. ¿Desea guardarlos?", parent=self)
            if respuesta is None or (respuesta and not self._guardar()):
                return
        self.destroy()