import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import pandas as pd
from pdf_engine import (
    PDFEngine,
    PDFPasswordRequiredError,
    PDFInvalidPasswordError
)
from PIL import Image, ImageTk
from version import __version__
from clasificador_conceptos import ClasificadorConceptos
from worker_manager import TrabajadorEnSegundoPlano
from ui_theme import (
    aplicar_estilos,
    construir_warning_box,
    COLOR_BG,
    COLOR_CANVAS_BG,
    FONT_BUTTON,
)

# Importación del backend de Banco de Bogotá
try:
    from script_ban_bogota import ejecutar_proceso_exportacion as exportar_bogota
    _import_error_msg = None
except Exception as e:
    exportar_bogota = None
    _import_error_msg = str(e)

try:
    from script import ejecutar_proceso_exportacion, reorganizar_excel
except ImportError as err:
    ejecutar_proceso_exportacion = None
    reorganizar_excel = None
    _import_error_msg = str(err)

try:
    from parser_excel import extraer_datos_desde_excel
except ImportError as err:
    extraer_datos_desde_excel = None
    _import_parser_error_msg = str(err)


class PDFViewerApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Lector de extractos v{__version__} - Tkinter")
        self.root.geometry("1100x720")
        self.root.minsize(900, 600)
        self.root.configure(bg=COLOR_BG)

        self.pdf_engine = PDFEngine()
        self.trabajador = TrabajadorEnSegundoPlano(self.root)

        # Mantiene vivas las referencias a los ImageTk.PhotoImage mostrados
        self._photo_images = []

        self._center_window()
        aplicar_estilos(self.root)
        self._build_ui()

    def _center_window(self):
        self.root.update_idletasks()
        w, h = 1100, 720
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------
    # Construcción de la interfaz
    # ------------------------------------------------------------------
    def _build_ui(self):
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # ============= PANEL IZQUIERDO =============
        left_frame = ttk.Frame(
            paned_window, width=270, padding=20, style="Sidebar.TFrame"
        )
        left_frame.pack_propagate(False)
        paned_window.add(left_frame, weight=1)

        base_dir = os.path.dirname(__file__)
        logo_path = os.path.join(base_dir, "assets", "img", "logo.png")

        if os.path.exists(logo_path):
            icon_image_pil = Image.open(logo_path).resize((24, 24))
            self.icon_title = ImageTk.PhotoImage(icon_image_pil)
            title_label = ttk.Label(
                left_frame,
                text=" Panel de Control",
                image=self.icon_title,
                compound=tk.LEFT,
                style="Title.TLabel",
            )
        else:
            title_label = ttk.Label(
                left_frame, text=" Panel de Control", style="Title.TLabel"
            )

        title_label.pack(anchor="w", padx=10, pady=5)

        subtitle_label = ttk.Label(
            left_frame,
            text="Extrae y organiza tus extractos",
            style="Subtitle.TLabel",
        )
        subtitle_label.pack(anchor="w", pady=(0, 18))

        ttk.Separator(left_frame, orient="horizontal").pack(
            fill=tk.X, pady=(0, 18)
        )

        # Bloque de selección de banco: primer paso obligatorio, muy visible.
        banco_frame = tk.Frame(
            left_frame,
            bg="#EAF1FB",
            highlightbackground="#2F6FED",
            highlightthickness=1,
            bd=0,
        )
        banco_frame.pack(fill=tk.X, pady=(0, 15))

        lbl_banco = tk.Label(
            banco_frame,
            text="  Selecciona tu banco",
            bg="#EAF1FB",
            fg="#2457BE",
            font=("Segoe UI", 10, "bold"),
        )
        lbl_banco.pack(anchor="w", padx=10, pady=(10, 4))

        self.combo_bancos = ttk.Combobox(
            banco_frame,
            values=[
                "-- Selecciona un banco --",
                "Bancolombia (Estándar)",
                "Banco de Bogotá (Extractos PyME)",
            ],
            state="readonly",
            font=("Segoe UI", 10),
        )
        self.combo_bancos.current(0)
        self.combo_bancos.pack(fill=tk.X, padx=10, pady=(0, 10))
        self.combo_bancos.bind("<<ComboboxSelected>>", self._on_banco_seleccionado)

        construir_warning_box(left_frame)

        # Botón para Extractos
        self.btn_extractos = ttk.Button(
            left_frame,
            text="📄 Cargar Extracto",
            command=lambda: self.cargar_y_procesar_pdf("Extracto"),
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_extractos.pack(fill=tk.X, pady=6)

        # Botón para Movimientos Persona Natural
        self.btn_movimientos_pn = ttk.Button(
            left_frame,
            text="📄 Movimientos Persona Natural",
            command=lambda: self.cargar_y_procesar_pdf("MovimientosPNatural"),
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_movimientos_pn.pack(fill=tk.X, pady=6)

        # Botón para Movimientos Sociedad
        self.btn_movimientos_soc = ttk.Button(
            left_frame,
            text="📄 Movimientos Sociedad",
            command=lambda: self.cargar_y_procesar_pdf("MovimientoSOC"),
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_movimientos_soc.pack(fill=tk.X, pady=6)

        self.btn_load_excel = ttk.Button(
            left_frame,
            text="📊 Cargar Excel",
            command=self.cargar_archivo_excel,
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_load_excel.pack(fill=tk.X, pady=6)

        self.btn_editarbancarios = ttk.Button(
            left_frame,
            text="✏️ Editar Bancarios",
            command=lambda: self.abrir_clasificador(),
            style="Primary.TButton",
            cursor="hand2",
            state=tk.DISABLED,
        )
        self.btn_editarbancarios.pack(fill=tk.X, pady=6)

        # Botones cuya disponibilidad depende del banco elegido arriba.
        self.botones_dependientes_banco = [
            self.btn_extractos,
            self.btn_movimientos_pn,
            self.btn_movimientos_soc,
            self.btn_load_excel,
            self.btn_editarbancarios,
        ]

        # Crear el botón "Limpiar / Nuevo Documento"
        self.btn_limpiar = tk.Button(
            left_frame,
            text="Limpiar Sesión",
            command=self.limpiar_sesion,
            bg="#f44336",
            fg="#FFFFFF",
            state="disabled",
        )
        self.btn_limpiar.pack(fill=tk.X, pady=6)

        self.btn_process = ttk.Button(
            left_frame,
            text="⚡  Extraer TXT y Excel",
            command=self.process_pdf,
            state=tk.DISABLED,
            style="Secondary.TButton",
            cursor="hand2",
        )
        self.btn_process.pack(fill=tk.X, pady=6)

        # ----------------------------------------------------------
        # BARRA DE PROGRESO
        # ----------------------------------------------------------
        progress_frame = ttk.Frame(left_frame, style="Sidebar.TFrame")
        progress_frame.pack(fill=tk.X, pady=(12, 4))

        self.progress_label = ttk.Label(
            progress_frame,
            text="Listo",
            style="Info.TLabel"
        )
        self.progress_label.pack(anchor="w", pady=(0, 4))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=0
        )
        self.progress_bar.pack(fill=tk.X)

        self.btn_open_excel = ttk.Button(
            left_frame,
            text="↗️ Abrir Excel",
            command=self.open_excel,
            state=tk.DISABLED,
            style="Secondary.TButton",
            cursor="hand2",
        )

        self.btn_open_excel.pack(fill=tk.X, pady=6)

        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=18)

        status_header = ttk.Label(
            left_frame, text="ESTADO", style="Subtitle.TLabel"
        )
        status_header.pack(anchor="w", pady=(0, 6))

        self.info_label = ttk.Label(
            left_frame,
            text="Ningún archivo cargado",
            wraplength=220,
            style="Info.TLabel",
            justify="left",
        )
        self.info_label.pack(anchor="w", fill=tk.X)

        spacer = ttk.Frame(left_frame, style="Sidebar.TFrame")
        spacer.pack(fill=tk.BOTH, expand=True)

        footer_label = ttk.Label(
            left_frame,
            text=f"Lector de extractos • v{__version__}",
            style="Footer.TLabel",
        )
        footer_label.pack(anchor="w", side=tk.BOTTOM, pady=(10, 0))

        # ============= PANEL DERECHO =============
        right_frame = ttk.Frame(paned_window, style="Content.TFrame")
        paned_window.add(right_frame, weight=4)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 1. Pestaña de Vista Previa PDF
        self.tab_pdf = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pdf, text="  📄 Vista Previa PDF  ")

        scrollbar = ttk.Scrollbar(self.tab_pdf, orient=tk.VERTICAL)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas = tk.Canvas(
            self.tab_pdf,
            bg=COLOR_CANVAS_BG,
            highlightthickness=0,
            yscrollcommand=scrollbar.set,
        )
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.canvas.yview)

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            ),
        )
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # 2. Pestaña de Excel Generado
        self.tab_excel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text="  📈 Vista Previa de Excel  ")

        self.notebook_hojas = ttk.Notebook(self.tab_excel)
        self.notebook_hojas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self._on_banco_seleccionado()

    # ------------------------------------------------------------------
    # PROGRESS BAR
    # ------------------------------------------------------------------

    def iniciar_progreso(self, texto="Procesando...", indeterminado=False):
        self.progress_label.config(text=texto)

        if indeterminado:
            self.progress_bar.config(mode="indeterminate")
            self.progress_bar.start(10)
        else:
            self.progress_bar.stop()
            self.progress_bar.config(
                mode="determinate",
                maximum=100,
                value=0
            )

        self.root.update_idletasks()

    def actualizar_progreso(self, valor, texto=None):
        if self.progress_bar["mode"] != "determinate":
            return

        valor = max(0, min(100, valor))
        self.progress_bar["value"] = valor

        if texto:
            self.progress_label.config(text=texto)

    def finalizar_progreso(self, texto="Proceso completado"):
        self.progress_bar.stop()
        self.progress_bar.config(
            mode="determinate",
            maximum=100,
            value=100
        )
        self.progress_label.config(text=texto)

    def resetear_progreso(self):
        self.progress_bar.stop()
        self.progress_bar.config(
            mode="determinate",
            maximum=100,
            value=0
        )
        self.progress_label.config(text="Listo")

    def _set_botones_bloqueados(self, bloqueados: bool):
        """Evita disparar una segunda tarea pesada mientras un hilo secundario trabaja."""
        estado = tk.DISABLED if bloqueados else tk.NORMAL
        for btn in (
            self.btn_process,
            self.btn_load_excel,
            self.btn_editarbancarios,
        ):
            btn.config(state=estado)

        for btn in self.botones_dependientes_banco:
            if btn is self.btn_load_excel or btn is self.btn_editarbancarios:
                continue
            if not bloqueados:
                # Al reactivar, respetar las reglas del banco seleccionado.
                self._on_banco_seleccionado()
                break
            btn.config(state=tk.DISABLED)

    # ------------------------------------------------------------------
    # Lógica y Eventos
    # ------------------------------------------------------------------

    def _on_banco_seleccionado(self, event=None):
        seleccion = self.combo_bancos.get()

        if seleccion.startswith("Bancolombia"):
            for btn in self.botones_dependientes_banco:
                btn.config(state=tk.NORMAL)
        elif seleccion.startswith("Banco de Bogotá"):
            for btn in self.botones_dependientes_banco:
                btn.config(state=tk.DISABLED)
            self.btn_extractos.config(state=tk.NORMAL)
        else:
            for btn in self.botones_dependientes_banco:
                btn.config(state=tk.DISABLED)

    def abrir_clasificador(self):
        ruta_excel = getattr(self.pdf_engine, "last_excel_path", None)

        if not ruta_excel:
            messagebox.showwarning(
                "Aviso",
                "Primero debes procesar un PDF para generar un archivo Excel."
            )
            return

        ClasificadorConceptos(self.root, ruta_excel, self.cargar_excel_en_gui)

    def cargar_archivo_excel(self):
        """Abre un explorador, selecciona un .xlsx y lo procesa con parser_excel.py

        en un hilo secundario para no congelar la interfaz.
        """
        if extraer_datos_desde_excel is None:
            messagebox.showerror(
                "Error de Módulo",
                f"No se pudo importar 'parser_excel.py':\n{_import_parser_error_msg}",
            )
            return

        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel a procesar",
            filetypes=[("Archivos de Excel", "*.xlsx;*.xls")],
        )

        if not file_path:
            return

        base_dir, file_name = os.path.split(file_path)
        name, ext = os.path.splitext(file_name)
        output_path = os.path.join(base_dir, f"{name}_procesado{ext}")

        self.root.config(cursor="watch")
        self._set_botones_bloqueados(True)
        self.iniciar_progreso(
            "Procesando y reorganizando Excel...", indeterminado=True
        )
        self.info_label.config(
            text="⏳ Procesando y reorganizando Excel...",
            style="Info.TLabel",
        )

        def tarea(cola):
            resultado = extraer_datos_desde_excel(file_path, output_path)
            cola.put(("exito", resultado, None))

        self.trabajador.ejecutar(
            tarea,
            on_exito=self._on_excel_procesado,
            on_error=self._on_error_excel_procesado,
        )

    def _on_excel_procesado(self, excel_procesado):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)

        if hasattr(self, "pdf_engine") and self.pdf_engine:
            self.pdf_engine.last_excel_path = excel_procesado

        self.finalizar_progreso("Excel procesado correctamente")

        filename_clean = os.path.basename(excel_procesado)
        self.info_label.config(
            text=f"📊 Excel procesado:\n{filename_clean}",
            style="Status.TLabel",
        )

        if hasattr(self, "btn_open_excel") and self.btn_open_excel:
            self.btn_open_excel.config(state=tk.NORMAL, style="Secondary.TButton")

        if hasattr(self, "btn_limpiar") and self.btn_limpiar:
            self.btn_limpiar.config(state="normal")

        messagebox.showinfo(
            "Proceso Exitoso",
            f"El archivo Excel se ha formateado correctamente.\n\nGuardado en:\n{excel_procesado}",
        )

    def _on_error_excel_procesado(self, error):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)
        self.info_label.config(
            text="⚠️ Error al procesar el Excel.",
            style="Info.TLabel",
        )
        messagebox.showerror(
            "Error de Procesamiento",
            f"No se pudo estructurar el archivo Excel:\n{str(error)}",
        )
        self.resetear_progreso()

    def cargar_y_procesar_pdf(self, prefijo_tipo):
        file_path = filedialog.askopenfilename(
            title=f"Seleccionar PDF de {prefijo_tipo}",
            filetypes=[("Archivos PDF", "*.pdf")],
        )

        if not file_path:
            return

        try:
            folder_path, old_filename = os.path.split(file_path)
            filename_lower = old_filename.lower()

            prefijo_formateado = f"{prefijo_tipo}_"
            etiquetas_existentes = [
                "extracto_",
                "movimiento_",
                "movimientos_",
                "movimientospnatural_",
                "movimientosoc_",
            ]

            ya_tiene_prefijo = any(
                filename_lower.startswith(tag) for tag in etiquetas_existentes
            )

            if not ya_tiene_prefijo:
                nuevo_nombre = f"{prefijo_formateado}{old_filename}"
                new_file_path = os.path.join(folder_path, nuevo_nombre)

                if file_path != new_file_path:
                    os.replace(file_path, new_file_path)
                    file_path = new_file_path

            self.current_pdf_path = file_path
            self.tipo_documento = prefijo_tipo
            self.load_pdf(file_path)

            if hasattr(self, "btn_limpiar") and self.btn_limpiar:
                try:
                    self.btn_limpiar.configure(state="normal")
                except Exception:
                    self.btn_limpiar.config(state="normal")

        except Exception as e:
            messagebox.showerror(
                "Error al renombrar",
                f"No se pudo asignar el nombre al archivo:\n{str(e)}",
            )

    def load_pdf(self, file_path):
        while True:
            try:
                total_pages = self.pdf_engine.open_pdf(file_path)
                break

            except PDFPasswordRequiredError:
                password = simpledialog.askstring(
                    "PDF protegido",
                    "Este PDF está protegido con contraseña.\n\n"
                    "Ingrese la contraseña para continuar:",
                    parent=self.root,
                    show="*"
                )

                if password is None:
                    return

                try:
                    total_pages = self.pdf_engine.open_pdf(
                        file_path, password=password
                    )
                    break

                except PDFInvalidPasswordError:
                    messagebox.showerror(
                        "Contraseña incorrecta",
                        "La contraseña ingresada no es correcta.\n\n"
                        "Inténtelo nuevamente.",
                        parent=self.root
                    )

                except Exception as e:
                    messagebox.showerror(
                        "Error de Carga",
                        f"No se pudo cargar el PDF:\n{str(e)}",
                        parent=self.root
                    )
                    return

            except Exception as e:
                messagebox.showerror(
                    "Error de Carga",
                    f"No se pudo cargar el PDF:\n{str(e)}",
                    parent=self.root
                )
                return

        # A partir de aquí el documento ya está abierto (operación rápida).
        # El renderizado de páginas (pesado) se ejecuta en un hilo secundario.
        for child in self.scrollable_frame.winfo_children():
            child.destroy()
        self._photo_images.clear()

        self.root.config(cursor="watch")
        self._set_botones_bloqueados(True)
        self.iniciar_progreso("Renderizando páginas... 0/%d" % total_pages)

        def tarea(cola):
            def callback(indice, total):
                porcentaje = (indice / total) * 100
                cola.put((
                    "progreso",
                    porcentaje,
                    f"Renderizando páginas... {indice}/{total}",
                ))

            self.pdf_engine.renderizar_paginas(progreso_callback=callback)
            cola.put(("exito", (file_path, total_pages), None))

        self.trabajador.ejecutar(
            tarea,
            on_progreso=self.actualizar_progreso,
            on_exito=lambda datos: self._mostrar_paginas_renderizadas(*datos),
            on_error=self._on_error_render_pdf,
        )

    def _mostrar_paginas_renderizadas(self, file_path, total_pages):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)

        for img in self.pdf_engine.page_images:
            photo = ImageTk.PhotoImage(img)
            self._photo_images.append(photo)

            lbl_page = ttk.Label(self.scrollable_frame, image=photo)
            lbl_page.pack(pady=10, padx=20)

        filename = os.path.basename(file_path)

        self.info_label.config(
            text=f"📄 {filename}\nTotal páginas: {total_pages}",
            style="Status.TLabel",
        )

        self.btn_process.config(state=tk.NORMAL, style="Primary.TButton")
        self.btn_open_excel.config(state=tk.DISABLED, style="Secondary.TButton")

        self.notebook.select(self.tab_pdf)
        self.finalizar_progreso(f"PDF cargado correctamente • {total_pages} páginas")

    def _on_error_render_pdf(self, error):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)
        self.resetear_progreso()
        messagebox.showerror(
            "Error de Carga",
            f"No se pudo renderizar el PDF:\n{str(error)}",
        )

    def process_pdf(self):
        if not self.pdf_engine.has_document:
            messagebox.showwarning("Atención", "Carga un archivo PDF primero.")
            return

        banco_seleccionado = (
            self.combo_bancos.get() if hasattr(self, "combo_bancos") else ""
        )
        usa_bogota = banco_seleccionado.startswith("Banco de Bogotá")

        if usa_bogota and exportar_bogota is None:
            messagebox.showerror(
                "Error de Módulo",
                f"No se pudo importar el script de Banco de Bogotá:\n{_import_error_msg}",
            )
            return

        if not usa_bogota and ejecutar_proceso_exportacion is None:
            messagebox.showerror(
                "Error de Módulo",
                f"No se pudo importar el script de Bancolombia:\n{_import_error_msg}",
            )
            return

        nombre_base_pdf = os.path.splitext(
            os.path.basename(self.pdf_engine.current_path)
        )[0]
        nombre_lower = nombre_base_pdf.lower()

        tipo_actual = getattr(self, "tipo_documento", "Extracto")
        etiqueta_sugerida = f"{tipo_actual}_"

        palabras_clave = [
            "extracto_",
            "extracto",
            "movimiento_",
            "movimientos_",
            "movimiento",
            "movimientos",
            "movimientospnatural_",
            "movimientospnatural",
            "movimientosoc_",
            "movimientosoc",
        ]
        ya_tiene_etiqueta = any(
            nombre_lower.startswith(p) for p in palabras_clave
        )

        if not ya_tiene_etiqueta:
            nombre_sugerido = f"{etiqueta_sugerida}{nombre_base_pdf}.xlsx"
        else:
            nombre_sugerido = f"{nombre_base_pdf}.xlsx"

        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo Excel como...",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")],
        )

        if not save_path:
            return

        self.root.config(cursor="watch")
        self._set_botones_bloqueados(True)
        self.iniciar_progreso("Procesando PDF, por favor espera...")
        self.info_label.config(
            text="⏳ Procesando PDF, por favor espera...",
            style="Info.TLabel"
        )

        pdf_path = self.pdf_engine.current_path

        def tarea(cola):
            def callback(indice, total):
                # La extracción se pondera hasta el 85%; el resto es la
                # generación/formato del Excel.
                porcentaje = (indice / total) * 85
                cola.put((
                    "progreso",
                    porcentaje,
                    f"Extrayendo movimientos... página {indice}/{total}",
                ))

            cola.put(("progreso", 0, "Leyendo el PDF..."))

            if usa_bogota:
                excel_generado = exportar_bogota(
                    pdf_path,
                    output_excel_path=save_path,
                    progreso_callback=callback,
                )
            else:
                excel_generado = ejecutar_proceso_exportacion(
                    pdf_path,
                    output_excel_path=save_path,
                    progreso_callback=callback,
                )

            if not excel_generado:
                raise ValueError("No se obtuvieron registros del documento procesado.")

            cola.put(("progreso", 95, "Generando archivo Excel..."))
            cola.put(("exito", excel_generado, None))

        self.trabajador.ejecutar(
            tarea,
            on_progreso=self.actualizar_progreso,
            on_exito=self._on_pdf_procesado,
            on_error=self._on_error_process_pdf,
        )

    def _on_pdf_procesado(self, excel_generado):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)

        self.pdf_engine.last_excel_path = excel_generado

        self.btn_open_excel.config(state=tk.NORMAL, style="Secondary.TButton")

        self.info_label.config(
            text=(
                "✅ Excel"
                f" generado:\n{os.path.basename(excel_generado)}"
            ),
            style="Status.TLabel",
        )

        ClasificadorConceptos(self.root, excel_generado, self.cargar_excel_en_gui)
        self.cargar_excel_en_gui(excel_generado)
        self.finalizar_progreso("Extracción completada correctamente")

        messagebox.showinfo(
            "Proceso Exitoso",
            "¡Extracción completada!\n\nArchivo guardado"
            f" en:\n{excel_generado}",
        )

    def _on_error_process_pdf(self, error):
        self.root.config(cursor="")
        self._set_botones_bloqueados(False)
        self.info_label.config(
            text="⚠️ Ocurrió un error al procesar el PDF.",
            style="Info.TLabel",
        )
        messagebox.showerror(
            "Error al procesar",
            f"Ocurrió un error en la extracción:\n{str(error)}",
        )
        self.resetear_progreso()

    def cargar_excel_en_gui(self, excel_path):
        if not excel_path or not os.path.exists(excel_path):
            return

        try:
            for tab in self.notebook_hojas.tabs():
                self.notebook_hojas.forget(tab)

            excel_file = pd.ExcelFile(excel_path)

            for sheet_name in excel_file.sheet_names:
                df = pd.read_excel(excel_file, sheet_name=sheet_name)

                tab_frame = ttk.Frame(self.notebook_hojas)
                self.notebook_hojas.add(tab_frame, text=f" 📄 {sheet_name} ")

                scroll_y = ttk.Scrollbar(tab_frame, orient=tk.VERTICAL)
                scroll_x = ttk.Scrollbar(tab_frame, orient=tk.HORIZONTAL)

                cols = list(df.columns)
                tree = ttk.Treeview(
                    tab_frame,
                    columns=cols,
                    show="headings",
                    yscrollcommand=scroll_y.set,
                    xscrollcommand=scroll_x.set,
                )

                scroll_y.config(command=tree.yview)
                scroll_x.config(command=tree.xview)

                scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
                scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
                tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

                for col in cols:
                    tree.heading(col, text=col)
                    tree.column(
                        col, width=max(120, len(str(col)) * 12), anchor="center"
                    )

                for _, row in df.iterrows():
                    valores_fila = [
                        "" if pd.isna(val) else str(val) for val in row.values
                    ]
                    tree.insert("", tk.END, values=valores_fila)

            self.notebook.select(self.tab_excel)

            if hasattr(self, "btn_limpiar") and self.btn_limpiar is not None:
                try:
                    self.btn_limpiar.configure(state="normal")
                except Exception:
                    self.btn_limpiar.config(state="normal")

        except Exception as e:
            messagebox.showwarning(
                "Aviso", f"No se pudo cargar la vista previa del Excel:\n{e}"
            )

    def liberar_pdf(self):
        try:
            if hasattr(self, "pdf_engine") and self.pdf_engine:
                if hasattr(self.pdf_engine, "current_path"):
                    self.pdf_engine.current_path = None
                if hasattr(self.pdf_engine, "last_pdf_path"):
                    self.pdf_engine.last_pdf_path = None

            if hasattr(self, "current_pdf_path"):
                self.current_pdf_path = None

            if hasattr(self, "scrollable_frame") and self.scrollable_frame:
                for child in self.scrollable_frame.winfo_children():
                    child.destroy()

            self._photo_images.clear()

            if hasattr(self, "canvas") and self.canvas:
                self.canvas.delete("all")
                if hasattr(self, "scrollable_frame"):
                    self.canvas_window = self.canvas.create_window(
                        (0, 0), window=self.scrollable_frame, anchor="nw"
                    )
                self.canvas.configure(scrollregion=(0, 0, 0, 0))
                self.canvas.yview_moveto(0)

            if hasattr(self, "pdf_images"):
                self.pdf_images = []
            if hasattr(self, "pdf_image"):
                self.pdf_image = None

            if hasattr(self, "lbl_pdf_nombre") and self.lbl_pdf_nombre:
                self.lbl_pdf_nombre.config(text="")

            if hasattr(self, "lbl_total_paginas") and self.lbl_total_paginas:
                self.lbl_total_paginas.config(text="")

            if hasattr(self, "info_label") and self.info_label:
                self.info_label.config(text="")

        except Exception as e:
            messagebox.showerror(
                "Error al liberar PDF",
                f"Ocurrió un error al limpiar la vista previa del PDF:\n{e}",
            )

    def liberar_excel(self):
        try:
            if hasattr(self, "pdf_engine") and self.pdf_engine:
                if hasattr(self.pdf_engine, "last_excel_path"):
                    self.pdf_engine.last_excel_path = None
                if hasattr(self.pdf_engine, "df_original_raw"):
                    self.pdf_engine.df_original_raw = None

            if hasattr(self, "notebook_hojas") and self.notebook_hojas:
                for tab in self.notebook_hojas.tabs():
                    self.notebook_hojas.forget(tab)

            if hasattr(self, "btn_open_excel") and self.btn_open_excel:
                try:
                    self.btn_open_excel.configure(state="disabled")
                except Exception:
                    self.btn_open_excel.config(state="disabled")

            if hasattr(self, "lbl_excel_estado") and self.lbl_excel_estado:
                self.lbl_excel_estado.config(text="")

        except Exception as e:
            messagebox.showerror(
                "Error al liberar Excel",
                f"Ocurrió un error al limpiar la vista previa del Excel:\n{e}",
            )

    def limpiar_sesion(self):
        try:
            self.liberar_pdf()
            self.liberar_excel()

            if hasattr(self, "btn_limpiar") and self.btn_limpiar:
                try:
                    self.btn_limpiar.configure(state="disabled")
                except Exception:
                    self.btn_limpiar.config(state="disabled")

            if hasattr(self, "btn_open_excel"):
                self.btn_open_excel.config(state="disabled")

            messagebox.showinfo(
                "Limpieza realizada",
                "Se ha restablecido la interfaz y se liberaron el PDF y el Excel.",
            )

        except Exception as e:
            messagebox.showerror(
                "Error", f"Ocurrió un error al limpiar la sesión:\n{e}"
            )

    def open_excel(self):
        if not self.pdf_engine.open_generated_excel():
            messagebox.showerror("Error", "No se encontró el archivo Excel.")

    def _on_mousewheel(self, event):
        if self.pdf_engine.has_document:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


if __name__ == "__main__":
    root = tk.Tk()
    base_dir = os.path.dirname(__file__)
    icon_path = os.path.join(base_dir, "assets", "img", "logo.png")
    try:
        icono = tk.PhotoImage(file=icon_path)
        root.iconphoto(True, icono)
    except tk.TclError:
        pass

    app = PDFViewerApp(root)
    root.mainloop()