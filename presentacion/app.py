import os
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox, simpledialog, ttk

import pandas as pd
from PIL import Image, ImageTk

from infraestructura.excel import abrir_con_sistema
from infraestructura.pdf_engine import PDFEngine, PDFInvalidPasswordError, PDFPasswordRequiredError
from nucleo.config import URL_MANUAL, ruta_asset
from presentacion import actualizaciones
from presentacion.componentes import elegir_opcion
from presentacion.dialogo_plano import generar_plano_con_dialogo
from presentacion.hilos import TrabajadorEnSegundoPlano
from presentacion.tema import COLOR_BG, COLOR_CANVAS_BG, aplicar_estilos, construir_warning_box
from servicios.fuentes import (BTN_EXCEL, BTN_EXTRACTO, BTN_PLANILLA, BTN_PLANO, BTN_PN, BTN_SOC, FUENTES,
                               SIN_FUENTE, SIN_REGISTROS, con_prefijo, procesar_excel_bancolombia)
from version import __version__

ANCHO, ALTO = 1100, 720


class PDFViewerApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Lector de extractos v{__version__} - Tkinter")
        self.root.minsize(ANCHO, ALTO)
        self.root.configure(bg=COLOR_BG)

        self.pdf_engine = PDFEngine()
        self.trabajador = TrabajadorEnSegundoPlano(self.root)
        self.tipo_documento = BTN_EXTRACTO
        self.ultimo_excel = None
        self.opciones_planilla = {}
        self._photo_images = []
        self._editor_adm = None

        self._center_window()
        aplicar_estilos(self.root)
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._cerrar)
        actualizaciones.programar(self.root, al_cerrar=self._cerrar)

    def _cerrar(self):
        try:
            self.pdf_engine.liberar_recursos()
        except Exception:
            pass
        self.root.destroy()

    def _center_window(self):
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - ANCHO) // 2
        y = (self.root.winfo_screenheight() - ALTO) // 2
        self.root.geometry(f"{ANCHO}x{ALTO}+{x}+{y}")

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    # ------------------------------------------------------------------
    def _boton(self, parent, texto, comando, estilo="Primary.TButton", estado=tk.DISABLED):
        btn = ttk.Button(parent, text=texto, command=comando, style=estilo, cursor="hand2", state=estado)
        btn.pack(fill=tk.X, pady=6)
        return btn

    def _build_ui(self):
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left_frame = ttk.Frame(paned_window, width=270, padding=20, style="Sidebar.TFrame")
        left_frame.pack_propagate(False)
        paned_window.add(left_frame, weight=1)

        logo_path = ruta_asset("img", "logo.png")
        self.icon_title = ImageTk.PhotoImage(Image.open(logo_path).resize((24, 24))) if os.path.exists(logo_path) else None
        ttk.Label(left_frame, text=" Panel de Control", image=self.icon_title or "", compound=tk.LEFT,
                  style="Title.TLabel").pack(anchor="w", padx=10, pady=5)
        ttk.Label(left_frame, text="Extrae y organiza tus extractos", style="Subtitle.TLabel").pack(anchor="w", pady=(0, 18))
        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 18))

        construir_warning_box(left_frame)
        banco_frame = tk.Frame(left_frame, bg="#EAF1FB", highlightbackground="#2F6FED", highlightthickness=1, bd=0)
        banco_frame.pack(fill=tk.X, pady=(0, 15))
        tk.Label(banco_frame, text="  Selecciona tu banco", bg="#EAF1FB", fg="#2457BE",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
        self.combo_bancos = ttk.Combobox(banco_frame, values=[SIN_FUENTE, *FUENTES], state="readonly",
                                         font=("Segoe UI", 10))
        self.combo_bancos.current(0)
        self.combo_bancos.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.combo_bancos.bind("<<ComboboxSelected>>", self._on_banco_seleccionado)

        cargar = lambda tipo: lambda: self.cargar_y_procesar_pdf(tipo)  # noqa: E731
        self.botones = {
            BTN_EXTRACTO: self._boton(left_frame, "📄 Cargar Extracto", cargar(BTN_EXTRACTO)),
            BTN_PLANILLA: self._boton(left_frame, "🧾 Cargar Planilla Seguridad Social", cargar(BTN_PLANILLA)),
            BTN_PN: self._boton(left_frame, "📄 Movimientos Persona Natural", cargar(BTN_PN)),
            BTN_SOC: self._boton(left_frame, "📄 Movimientos Sociedad", cargar(BTN_SOC)),
            BTN_EXCEL: self._boton(left_frame, "📊 Cargar Excel", self.cargar_archivo_excel),
            BTN_PLANO: self._boton(left_frame, "🗄️ Generar Archivo Plano", self.generar_archivo_plano_manual),
        }

        self.btn_limpiar = tk.Button(left_frame, text="Limpiar Sesión", command=self.limpiar_sesion,
                                     bg="#f44336", fg="#FFFFFF", state="disabled")
        self.btn_limpiar.pack(fill=tk.X, pady=6)
        self.btn_process = self._boton(left_frame, "⚡  Extraer TXT y Excel", self.process_pdf, "Secondary.TButton")

        progress_frame = ttk.Frame(left_frame, style="Sidebar.TFrame")
        progress_frame.pack(fill=tk.X, pady=(12, 4))
        self.progress_label = ttk.Label(progress_frame, text="Listo", style="Info.TLabel")
        self.progress_label.pack(anchor="w", pady=(0, 4))
        self.progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", maximum=100, value=0)
        self.progress_bar.pack(fill=tk.X)

        self.btn_open_excel = self._boton(left_frame, "↗️ Abrir Excel", self.open_excel, "Secondary.TButton")
        self._boton(left_frame, "🏛️ Administradoras", self._abrir_administradoras, "Secondary.TButton", tk.NORMAL)
        self._boton(left_frame, "⬆ Actualizaciones", lambda: actualizaciones.verificar(self.root, True, self._cerrar),
                    "Secondary.TButton", tk.NORMAL)
        self._boton(left_frame, "Abrir manual 📒", lambda: webbrowser.open(URL_MANUAL), "Secondary.TButton", tk.NORMAL)

        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=18)
        ttk.Label(left_frame, text="ESTADO", style="Subtitle.TLabel").pack(anchor="w", pady=(0, 6))
        self.info_label = ttk.Label(left_frame, text="Ningún archivo cargado", wraplength=220,
                                    style="Info.TLabel", justify="left")
        self.info_label.pack(anchor="w", fill=tk.X)
        ttk.Frame(left_frame, style="Sidebar.TFrame").pack(fill=tk.BOTH, expand=True)
        ttk.Label(left_frame, text=f"Lector de extractos • v{__version__}",
                  style="Footer.TLabel").pack(anchor="w", side=tk.BOTTOM, pady=(10, 0))

        right_frame = ttk.Frame(paned_window, style="Content.TFrame")
        paned_window.add(right_frame, weight=4)
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_pdf = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pdf, text="  📄 Vista Previa PDF  ")
        scrollbar = ttk.Scrollbar(self.tab_pdf, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas = tk.Canvas(self.tab_pdf, bg=COLOR_CANVAS_BG, highlightthickness=0, yscrollcommand=scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.scrollable_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.tab_excel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text="  📈 Vista Previa de Excel  ")
        self.notebook_hojas = ttk.Notebook(self.tab_excel)
        self.notebook_hojas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self._on_banco_seleccionado()

    # ------------------------------------------------------------------
    # Progreso y estado
    # ------------------------------------------------------------------
    def iniciar_progreso(self, texto="Procesando...", indeterminado=False):
        self.progress_label.config(text=texto)
        if indeterminado:
            self.progress_bar.config(mode="indeterminate")
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.progress_bar.config(mode="determinate", maximum=100, value=0)
        self.root.update_idletasks()

    def actualizar_progreso(self, valor, texto=None):
        if self.progress_bar["mode"] != "determinate":
            return
        self.progress_bar["value"] = max(0, min(100, valor))
        if texto:
            self.progress_label.config(text=texto)

    def finalizar_progreso(self, texto="Proceso completado", valor=100):
        self.progress_bar.stop()
        self.progress_bar.config(mode="determinate", maximum=100, value=valor)
        self.progress_label.config(text=texto)

    def resetear_progreso(self):
        self.finalizar_progreso("Listo", 0)

    def _estado(self, texto, ok=False):
        self.info_label.config(text=texto, style="Status.TLabel" if ok else "Info.TLabel")

    def _iniciar_tarea(self, texto, indeterminado=False):
        self.root.config(cursor="watch")
        self._set_botones_bloqueados(True)
        self.iniciar_progreso(texto, indeterminado)
        self._estado(f"⏳ {texto}")

    def _terminar_tarea(self):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)

    def _error_tarea(self, titulo, estado, prefijo):
        def manejar(error):
            self._terminar_tarea()
            self._estado(f"⚠️ {estado}")
            messagebox.showerror(titulo, f"{prefijo}\n{error}")
            self.resetear_progreso()
        return manejar

    def _set_botones_bloqueados(self, bloqueados: bool):
        """Evita disparar una segunda tarea pesada mientras un hilo secundario trabaja."""
        if bloqueados:
            for btn in (*self.botones.values(), self.btn_process):
                btn.config(state=tk.DISABLED)
            return
        self._on_banco_seleccionado()
        self.btn_process.config(state=tk.NORMAL if self.pdf_engine.has_document else tk.DISABLED)

    def _registrar_excel(self, ruta):
        self.ultimo_excel = ruta
        self.btn_open_excel.config(state=tk.NORMAL, style="Secondary.TButton")
        self.btn_limpiar.config(state="normal")

    # ------------------------------------------------------------------
    # Selección de fuente
    # ------------------------------------------------------------------
    def _fuente(self):
        return FUENTES.get(self.combo_bancos.get())

    def _modulo_disponible(self, fuente):
        try:
            fuente.modulo()
            return True
        except ImportError as e:
            messagebox.showerror("Error de Módulo", str(e))
            return False

    def _on_banco_seleccionado(self, event=None):
        habilitados = self._fuente().botones if self._fuente() else ()
        for clave, btn in self.botones.items():
            btn.config(state=tk.NORMAL if clave in habilitados else tk.DISABLED)

    # ------------------------------------------------------------------
    # Excel del banco y plano manual
    # ------------------------------------------------------------------
    def generar_archivo_plano_manual(self):
        file_path = filedialog.askopenfilename(title="Seleccionar Excel procesado para Archivo Plano",
                                               filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
        if file_path:
            generar_plano_con_dialogo(self.root, file_path)

    def cargar_archivo_excel(self):
        file_path = filedialog.askopenfilename(title="Seleccionar archivo Excel a procesar",
                                               filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
        if not file_path:
            return
        self._iniciar_tarea("Procesando y reorganizando Excel...", indeterminado=True)
        self.trabajador.ejecutar_funcion(
            lambda: procesar_excel_bancolombia(file_path),
            on_exito=self._on_excel_procesado,
            on_error=self._error_tarea("Error de Procesamiento", "Error al procesar el Excel.",
                                       "No se pudo estructurar el archivo Excel:"),
        )

    def _on_excel_procesado(self, excel_procesado):
        self._terminar_tarea()
        self._registrar_excel(excel_procesado)
        self.cargar_excel_en_gui(excel_procesado)
        self.finalizar_progreso("Excel procesado correctamente")
        self._estado(f"📊 Excel procesado:\n{os.path.basename(excel_procesado)}", ok=True)
        messagebox.showinfo("Proceso Exitoso",
                            f"El archivo Excel se ha formateado correctamente.\n\nGuardado en:\n{excel_procesado}")

    # ------------------------------------------------------------------
    # Carga de documentos
    # ------------------------------------------------------------------
    def cargar_y_procesar_pdf(self, prefijo_tipo):
        fuente = self._fuente()
        if not fuente:
            return
        es_ss = prefijo_tipo == BTN_PLANILLA
        if es_ss:
            if not self._modulo_disponible(fuente):
                return
            self.opciones_planilla = self._pedir_opciones(fuente)
            if self.opciones_planilla is None:
                return
        acepta_excel = es_ss and fuente.acepta_excel
        titulo = f"Seleccionar {'PDF o Excel' if acepta_excel else 'PDF'} de {prefijo_tipo} (puede elegir varias)"
        tipos = [("PDF o Excel", "*.pdf;*.xlsx;*.xls")] if acepta_excel else [("Archivos PDF", "*.pdf")]
        rutas = list(filedialog.askopenfilenames(title=titulo, filetypes=tipos))
        if not rutas:
            return
        try:
            rutas = [con_prefijo(r, prefijo_tipo) for r in rutas]
        except Exception as e:
            messagebox.showerror("Error al renombrar", f"No se pudo asignar el nombre al archivo:\n{e}")
            return

        self.tipo_documento = prefijo_tipo
        if es_ss and (len(rutas) > 1 or not rutas[0].lower().endswith(".pdf")):
            self._procesar_seguridad_social(rutas)
        elif len(rutas) > 1:
            self._procesar_extractos_multiples(rutas)
        else:
            self.load_pdf(rutas[0])
            self.btn_limpiar.config(state="normal")

    def load_pdf(self, file_path):
        password = None
        while True:
            try:
                total_pages = self.pdf_engine.open_pdf(file_path, password=password)
                break
            except PDFPasswordRequiredError:
                pass
            except PDFInvalidPasswordError:
                messagebox.showerror("Contraseña incorrecta",
                                     "La contraseña ingresada no es correcta.\n\nInténtelo nuevamente.", parent=self.root)
            except Exception as e:
                messagebox.showerror("Error de Carga", f"No se pudo cargar el PDF:\n{e}", parent=self.root)
                return
            password = simpledialog.askstring(
                "PDF protegido", "Este PDF está protegido con contraseña.\n\nIngrese la contraseña para continuar:",
                parent=self.root, show="*")
            if password is None:
                return

        self._limpiar_vista_pdf()
        self._iniciar_tarea(f"Renderizando páginas... 0/{total_pages}")

        def tarea(cola):
            self.pdf_engine.renderizar_paginas(progreso_callback=lambda i, t: cola.put(
                ("progreso", i / t * 100, f"Renderizando páginas... {i}/{t}")))
            cola.put(("exito", (file_path, total_pages), None))

        self.trabajador.ejecutar(
            tarea,
            on_progreso=self.actualizar_progreso,
            on_exito=lambda datos: self._mostrar_paginas_renderizadas(*datos),
            on_error=self._error_tarea("Error de Carga", "No se pudo renderizar el PDF.", "No se pudo renderizar el PDF:"),
        )

    def _mostrar_paginas_renderizadas(self, file_path, total_pages):
        self._terminar_tarea()
        for img in self.pdf_engine.page_images:
            photo = ImageTk.PhotoImage(img)
            self._photo_images.append(photo)
            ttk.Label(self.scrollable_frame, image=photo).pack(pady=10, padx=20)
        self._estado(f"📄 {os.path.basename(file_path)}\nTotal páginas: {total_pages}", ok=True)
        self.btn_process.config(state=tk.NORMAL, style="Primary.TButton")
        self.btn_open_excel.config(state=tk.DISABLED, style="Secondary.TButton")
        self.notebook.select(self.tab_pdf)
        self.finalizar_progreso(f"PDF cargado correctamente • {total_pages} páginas")

    # ------------------------------------------------------------------
    # Extracción
    # ------------------------------------------------------------------
    def _pedir_opciones(self, fuente, ruta=None):
        """Solicita las opciones que declara el extractor (p. ej. tipo de cuenta Davivienda). None = cancelado."""
        elegidas = {}
        for parametro, (titulo, texto, opciones, detectar) in fuente.opciones().items():
            try:
                sugerida = detectar(ruta) if detectar and ruta else None
            except Exception:
                sugerida = None
            valor = elegir_opcion(self.root, titulo, texto, opciones, sugerida)
            if not valor:
                return None
            elegidas[parametro] = valor
        return elegidas

    def _nombre_sugerido(self, ruta):
        nombre = os.path.splitext(os.path.basename(ruta))[0]
        if nombre.lower().startswith(("extracto", "movimiento")):
            return f"{nombre}.xlsx"
        return f"{self.tipo_documento}_{nombre}.xlsx"

    def process_pdf(self):
        if not self.pdf_engine.has_document:
            messagebox.showwarning("Atención", "Carga un archivo PDF primero.")
            return
        fuente = self._fuente()
        if not fuente or not self._modulo_disponible(fuente):
            return
        if fuente.planilla:
            return self._procesar_seguridad_social()

        pdf_path = self.pdf_engine.current_path
        opciones = self._pedir_opciones(fuente, pdf_path)
        if opciones is None:
            return
        save_path = filedialog.asksaveasfilename(title="Guardar archivo Excel como...",
                                                 initialfile=self._nombre_sugerido(pdf_path),
                                                 defaultextension=".xlsx", filetypes=[("Libro de Excel", "*.xlsx")])
        if not save_path:
            return

        self._iniciar_tarea("Procesando PDF, por favor espera...")
        tipo = self.tipo_documento

        def tarea(cola):
            cola.put(("progreso", 0, "Leyendo el PDF..."))
            excel = fuente.exportar(
                pdf_path, save_path, tipo_documento=tipo, **opciones,
                progreso=lambda i, t: cola.put(("progreso", i / t * 85, f"Extrayendo movimientos... página {i}/{t}")),
            )
            cola.put(("progreso", 95, "Generando archivo Excel..."))
            cola.put(("exito", excel, None))

        self.trabajador.ejecutar(
            tarea,
            on_progreso=self.actualizar_progreso,
            on_exito=self._on_pdf_procesado,
            on_error=self._error_tarea("Error al procesar", "Ocurrió un error al procesar el PDF.",
                                       "Ocurrió un error en la extracción:"),
        )

    def _on_pdf_procesado(self, excel_generado):
        self._terminar_tarea()
        self._registrar_excel(excel_generado)
        self._estado(f"✅ Excel generado:\n{os.path.basename(excel_generado)}", ok=True)
        self.cargar_excel_en_gui(excel_generado)
        self.finalizar_progreso("Extracción completada correctamente")
        messagebox.showinfo("Proceso Exitoso", f"¡Extracción completada!\n\nArchivo guardado en:\n{excel_generado}")
        generar_plano_con_dialogo(self.root, excel_generado)

    def _procesar_extractos_multiples(self, rutas):
        fuente = self._fuente()
        if not self._modulo_disponible(fuente):
            return
        tipo = self.tipo_documento
        self._iniciar_tarea(f"Procesando 0/{len(rutas)} archivos...")

        def tarea(cola):
            excels = []
            for i, ruta in enumerate(rutas):
                nombre = os.path.basename(ruta)
                cola.put(("progreso", i / len(rutas) * 100, f"Procesando {i + 1}/{len(rutas)}: {nombre}"))
                try:
                    excels.append(fuente.exportar(ruta, f"{os.path.splitext(ruta)[0]}.xlsx", tipo_documento=tipo))
                except Exception as e:
                    raise ValueError(f"{nombre}:\n{e or SIN_REGISTROS}") from e
            cola.put(("exito", excels, None))

        self.trabajador.ejecutar(
            tarea,
            on_progreso=self.actualizar_progreso,
            on_exito=self._on_extractos_multiples,
            on_error=self._error_tarea("Error al procesar", "Ocurrió un error al procesar el PDF.",
                                       "Ocurrió un error en la extracción:"),
        )

    def _on_extractos_multiples(self, excels):
        from servicios.plano import GeneradorExtractosMultiples

        self._terminar_tarea()
        self._registrar_excel(excels[-1])
        self.cargar_excel_en_gui(excels[-1])
        self.finalizar_progreso(f"{len(excels)} extractos procesados")
        nombres = "\n".join(os.path.basename(e) for e in excels)
        self._estado(f"✅ Excel generados:\n{nombres}", ok=True)
        messagebox.showinfo("Proceso Exitoso", f"¡Extracción completada!\n\nArchivos generados:\n{nombres}")
        generar_plano_con_dialogo(self.root, excels[0], generador=GeneradorExtractosMultiples(excels), pedir_ruta=True)

    def _procesar_seguridad_social(self, rutas=None):
        fuente = self._fuente()
        if not self._modulo_disponible(fuente):
            return
        from servicios.seguridad_social import generador_para

        rutas = rutas or [self.pdf_engine.current_path]
        generador = generador_para(fuente.generador(), rutas, **self.opciones_planilla)
        ruta = generar_plano_con_dialogo(self.root, rutas[0], generador=generador, confirmar=False, pedir_ruta=True)
        if ruta:
            self._registrar_excel(ruta)
            self._estado(f"✅ Plano generado:\n{os.path.basename(ruta)}", ok=True)

    # ------------------------------------------------------------------
    # Vista previa de Excel
    # ------------------------------------------------------------------
    def cargar_excel_en_gui(self, excel_path):
        if not excel_path or not os.path.exists(excel_path):
            return
        try:
            self._limpiar_vista_excel()
            with pd.ExcelFile(excel_path) as excel_file:
                hojas = {h: pd.read_excel(excel_file, sheet_name=h) for h in excel_file.sheet_names}
            for sheet_name, df in hojas.items():
                tab_frame = ttk.Frame(self.notebook_hojas)
                self.notebook_hojas.add(tab_frame, text=f" 📄 {sheet_name} ")
                scroll_y = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL)
                scroll_x = ttk.Scrollbar(tab_frame, orient=tk.HORIZONTAL)
                cols = [str(c) for c in df.columns]
                tree = ttk.Treeview(tab_frame, columns=cols, show="headings",
                                    yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
                scroll_y.config(command=tree.yview)
                scroll_x.config(command=tree.xview)
                scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
                scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
                tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                for col in cols:
                    tree.heading(col, text=col)
                    tree.column(col, width=max(120, len(col) * 12), anchor="center")
                for fila in df.astype(object).where(df.notna(), "").itertuples(index=False):
                    tree.insert("", tk.END, values=[str(v) for v in fila])
            self.notebook.select(self.tab_excel)
            self.btn_limpiar.config(state="normal")
        except Exception as e:
            messagebox.showwarning("Aviso", f"No se pudo cargar la vista previa del Excel:\n{e}")

    # ------------------------------------------------------------------
    # Sesión
    # ------------------------------------------------------------------
    def _limpiar_vista_pdf(self):
        for child in self.scrollable_frame.winfo_children():
            child.destroy()
        self._photo_images.clear()
        self.canvas.configure(scrollregion=(0, 0, 0, 0))
        self.canvas.yview_moveto(0)

    def _limpiar_vista_excel(self):
        for tab in self.notebook_hojas.tabs():
            self.notebook_hojas.forget(tab)

    def limpiar_sesion(self):
        try:
            self.pdf_engine.liberar_recursos()
            self._limpiar_vista_pdf()
            self._limpiar_vista_excel()
            self.ultimo_excel = None
            self._estado("")
            self.resetear_progreso()
            for btn in (self.btn_limpiar, self.btn_open_excel, self.btn_process):
                btn.config(state="disabled")
            messagebox.showinfo("Limpieza realizada", "Se ha restablecido la interfaz y se liberaron el PDF y el Excel.")
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al limpiar la sesión:\n{e}")

    def open_excel(self):
        if not abrir_con_sistema(self.ultimo_excel):
            messagebox.showerror("Error", "No se encontró el archivo Excel.")

    def _abrir_administradoras(self):
        if self._editor_adm is not None and self._editor_adm.winfo_exists():
            self._editor_adm.deiconify()
            self._editor_adm.lift()
            self._editor_adm.focus_force()
            return
        try:
            from presentacion.editor_administradoras import EditorAdministradoras
            self._editor_adm = EditorAdministradoras(self.root)
        except Exception as e:
            messagebox.showerror("Administradoras", f"No se pudo abrir administradoras.json:\n{e}")

    def _on_mousewheel(self, event):
        if self.pdf_engine.has_document:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def main():
    root = tk.Tk()
    try:
        root.iconphoto(True, tk.PhotoImage(file=ruta_asset("img", "logo.png")))
    except tk.TclError:
        pass
    PDFViewerApp(root)
    root.mainloop()