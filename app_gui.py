import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from pdf_engine import PDFEngine
from PIL import Image, ImageTk
from version import __version__

# ---- AGREGA ESTA LÍNEA ----
from clasificador_conceptos import ClasificadorConceptos 
# ---------------------------

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

# ------------------------------------------------------------------
# Paleta y constantes visuales
# ------------------------------------------------------------------
COLOR_BG = "#F4F6F8"
COLOR_SIDEBAR = "#FFFFFF"
COLOR_PRIMARY = "#2F6FED"
COLOR_PRIMARY_DARK = "#2457BE"
COLOR_TEXT = "#1F2933"
COLOR_MUTED = "#7B8794"
COLOR_SUCCESS = "#2E7D32"
COLOR_BORDER = "#E4E7EB"
COLOR_CANVAS_BG = "#525659"

FONT_TITLE = ("Segoe UI", 16, "bold")
FONT_SUBTITLE = ("Segoe UI", 9)
FONT_BUTTON = ("Segoe UI", 10)
FONT_INFO = ("Segoe UI", 9)
FONT_STATUS = ("Segoe UI", 9, "bold")


class PDFViewerApp:

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Lector de extractos v{__version__} - Tkinter")
        self.root.geometry("1100x720")
        self.root.minsize(900, 600)
        self.root.configure(bg=COLOR_BG)

        self.pdf_engine = PDFEngine()

        self._center_window()
        self._apply_styles()
        self._build_ui()

    def _center_window(self):
        self.root.update_idletasks()
        w, h = 1100, 720
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _apply_styles(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=COLOR_BG)
        style.configure("Sidebar.TFrame", background=COLOR_SIDEBAR)
        style.configure("Content.TFrame", background=COLOR_BG)

        style.configure(
            "Title.TLabel",
            background=COLOR_SIDEBAR,
            foreground=COLOR_TEXT,
            font=FONT_TITLE,
        )
        style.configure(
            "Subtitle.TLabel",
            background=COLOR_SIDEBAR,
            foreground=COLOR_MUTED,
            font=FONT_SUBTITLE,
        )
        style.configure(
            "Info.TLabel",
            background=COLOR_SIDEBAR,
            foreground=COLOR_MUTED,
            font=FONT_INFO,
        )
        style.configure(
            "Status.TLabel",
            background=COLOR_SIDEBAR,
            foreground=COLOR_SUCCESS,
            font=FONT_STATUS,
        )
        style.configure(
            "Footer.TLabel",
            background=COLOR_SIDEBAR,
            foreground=COLOR_MUTED,
            font=("Segoe UI", 8),
        )

        style.configure(
            "Primary.TButton",
            font=FONT_BUTTON,
            padding=10,
            background=COLOR_PRIMARY,
            foreground="white",
            borderwidth=0,
            focusthickness=0,
        )
        style.map(
            "Primary.TButton",
            background=[
                ("active", COLOR_PRIMARY_DARK),
                ("disabled", "#B9C6E4"),
            ],
            foreground=[("disabled", "#F0F0F0")],
        )

        style.configure(
            "Secondary.TButton",
            font=FONT_BUTTON,
            padding=10,
            background="#EDF1F7",
            foreground=COLOR_TEXT,
            borderwidth=0,
        )
        style.map(
            "Secondary.TButton",
            background=[("active", "#DDE4EF"), ("disabled", "#F3F4F6")],
            foreground=[("disabled", "#B0B7C3")],
        )

        style.configure("TPanedwindow", background=COLOR_BG)
        style.configure("TNotebook", background=COLOR_BG, borderwidth=0)
        style.configure("TNotebook.Tab", font=FONT_BUTTON, padding=(16, 8))
        style.map(
            "TNotebook.Tab",
            background=[("selected", "white")],
            foreground=[("selected", COLOR_PRIMARY)],
        )

        style.configure(
            "Treeview",
            font=("Segoe UI", 9),
            rowheight=26,
            background="white",
            fieldbackground="white",
            foreground=COLOR_TEXT,
        )
        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            background="#EEF1F5",
            foreground=COLOR_TEXT,
            padding=6,
        )
        style.map(
            "Treeview",
            background=[("selected", COLOR_PRIMARY)],
            foreground=[("selected", "white")],
        )

    def _build_warning_box(self, parent):
        warning_frame = tk.Frame(
            parent,
            bg="#FDECEA",
            highlightbackground="#D93025",
            highlightthickness=2,
            bd=0,
        )
        warning_frame.pack(fill=tk.X, pady=(0, 15))

        warning_label = tk.Label(
            warning_frame,
            text=(
                "⚠️ Recuerde que este programa únicamente puede usarse con"
                " archivos PDF que contengan texto (no imágenes escaneadas), y"
                " que el soporte actual es exclusivo para extractos con el"
                " formato de Bancolombia."
            ),
            bg="#FDECEA",
            fg="#D93025",
            font=("Segoe UI", 8),
            wraplength=220,
            justify="left",
            padx=10,
            pady=8,
        )
        warning_label.pack(fill=tk.X)

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

        
        self._build_warning_box(left_frame)
# Botón para Extractos
        self.btn_extractos = ttk.Button(
            left_frame,
            text="📄 Cargar Extracto",
            command=lambda: self.cargar_y_procesar_pdf("Extracto"),
            style="Primary.TButton",
            cursor="hand2",
        )
        self.btn_extractos.pack(fill=tk.X, pady=6)

        # Botón para Movimientos Persona Natural
        self.btn_movimientos_pn = ttk.Button(
            left_frame,
            text="📄 Movimientos Persona Natural",
            command=lambda: self.cargar_y_procesar_pdf("MovimientosPNatural"),
            style="Primary.TButton",
            cursor="hand2",
        )
        self.btn_movimientos_pn.pack(fill=tk.X, pady=6)

        # Botón para Movimientos Sociedad
        self.btn_movimientos_soc = ttk.Button(
            left_frame,
            text="📄 Movimientos Sociedad",
            command=lambda: self.cargar_y_procesar_pdf("MovimientoSOC"),
            style="Primary.TButton",
            cursor="hand2",
        )
        self.btn_movimientos_soc.pack(fill=tk.X, pady=6)
        btn_load_excel = ttk.Button(
                left_frame,
                text="📊 Cargar Excel",
                command=self.cargar_archivo_excel,  # <--- Vinculación agregada
                style="Primary.TButton",
                cursor="hand2",
            )
        btn_load_excel.pack(fill=tk.X, pady=6)

        btn_editarbancarios = ttk.Button(
            left_frame,
            text="✏️ Editar Bancarios",
            command=lambda: self.abrir_clasificador(),  # <--- Usamos un método intermediario o lambda
            style="Primary.TButton",
            cursor="hand2",
        )
        btn_editarbancarios.pack(fill=tk.X, pady=6)

# Crear el botón "Limpiar / Nuevo Documento"
        self.btn_limpiar = tk.Button(
            left_frame,  # o el frame/contenedor donde tengas tus botones
            text="Limpiar Sesión",
            command=self.limpiar_sesion,
            bg="#f44336",  # Color rojo/alerta (opcional)
            fg="#FFFFFF",
            state="disabled",
            
              # Empieza deshabilitado hasta que se cargue un archivo
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

        # Scrollbar y Canvas empaquetados explícitamente
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

        # Frame contenedor interno para las páginas del PDF
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.canvas_window = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )

        # Eventos para ajustar el scrollbar y permitir la rueda del ratón
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
    

        # ------------------------------------------------------------------
    # Lógica y Eventos
    # ------------------------------------------------------------------

    def abrir_clasificador(self):
        # Verificamos si ya hay un Excel generado y cargado en el sistema
        ruta_excel = getattr(self.pdf_engine, "last_excel_path", None)
        
        if not ruta_excel:
            messagebox.showwarning(
                "Aviso", 
                "Primero debes procesar un PDF para generar un archivo Excel."
            )
            return
            
        # Si existe, abrimos el clasificador pasándole la ruta actual
        ClasificadorConceptos(self.root, ruta_excel, self.cargar_excel_en_gui)
        
    def cargar_archivo_excel(self):
        """Abre un explorador para seleccionar un archivo .xlsx, lo procesa con

        parser_excel.py y muestra el resultado filtrado en la interfaz.
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

        try:
            self.root.config(cursor="watch")
            self.info_label.config(
                text="⏳ Procesando y reorganizando Excel...",
                style="Info.TLabel",
            )
            self.root.update_idletasks()

            # Definir la ruta de salida para el Excel procesado
            base_dir, file_name = os.path.split(file_path)
            name, ext = os.path.splitext(file_name)
            output_path = os.path.join(base_dir, f"{name}_procesado{ext}")

            # 1. Llamar a la función del nuevo archivo 'parser_excel.py'
            excel_procesado = extraer_datos_desde_excel(
                file_path, output_path
            )

            # 2. Guardar la ruta en el motor de la app
            if hasattr(self, "pdf_engine") and self.pdf_engine:
                self.pdf_engine.last_excel_path = excel_procesado

            # 3. Renderizar las pestañas (Movimientos y Resumen) en el Treeview
            self.cargar_excel_en_gui(excel_procesado)

            # 4. Actualizar estado y botones de la interfaz
            filename_clean = os.path.basename(excel_procesado)
            self.info_label.config(
                text=f"📊 Excel procesado:\n{filename_clean}",
                style="Status.TLabel",
            )

            if hasattr(self, "btn_open_excel") and self.btn_open_excel:
                self.btn_open_excel.config(
                    state=tk.NORMAL, style="Secondary.TButton"
                )

            if hasattr(self, "btn_limpiar") and self.btn_limpiar:
                self.btn_limpiar.config(state="normal")

            messagebox.showinfo(
                "Proceso Exitoso",
                f"El archivo Excel se ha formateado correctamente.\n\nGuardado en:\n{excel_procesado}",
            )

        except Exception as e:
            self.info_label.config(
                text="⚠️ Error al procesar el Excel.",
                style="Info.TLabel",
            )
            messagebox.showerror(
                "Error de Procesamiento",
                f"No se pudo estructurar el archivo Excel:\n{str(e)}",
            )
        finally:
            self.root.config(cursor="")

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

            # Lista de prefijos conocidos para evitar la duplicación de nombres
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

            # Guardar ruta y tipo actual
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
        try:
            total_pages = self.pdf_engine.open_pdf(file_path)

            # Limpiar páginas anteriores
            for child in self.scrollable_frame.winfo_children():
                child.destroy()

            # Renderizar las páginas en imágenes
            for photo in self.pdf_engine.page_images:
                lbl_page = ttk.Label(self.scrollable_frame, image=photo)
                lbl_page.pack(pady=10, padx=20)

            filename = os.path.basename(file_path)
            self.info_label.config(
                text=f"📄 {filename}\nTotal páginas: {total_pages}",
                style="Status.TLabel",
            )
            self.btn_process.config(state=tk.NORMAL, style="Primary.TButton")
           
            self.btn_open_excel.config(
                state=tk.DISABLED, style="Secondary.TButton"
            )

            # Cambiar a la vista previa del PDF
            self.notebook.select(self.tab_pdf)

        except Exception as e:
            messagebox.showerror(
                "Error de Carga", f"No se pudo cargar el PDF:\n{str(e)}"    
            )
    def generar_nombre_limpio(
        ruta_original, tipo_documento="Extracto_", extension_salida=".xlsx"
    ):
        """Genera un nombre de archivo asegurando que no se repitan los prefijos

        o sufijos como 'Extracto_' o 'Movimiento_'.
        """
        directorio, nombre_archivo = os.path.split(ruta_original)
        nombre_base, _ = os.path.splitext(nombre_archivo)

        # Convertimos a minúsculas solo para validar la existencia de palabras clave
        nombre_lower = nombre_base.lower()

        # Palabras clave que queremos evitar duplicar
        palabras_clave = [
            "extracto_",
            "extracto",
            "movimiento_",
            "movimientos_",
            "movimiento",
            "movimientos",
        ]

        # Verificar si el archivo YA contiene alguna de estas etiquetas al inicio o final
        ya_tiene_etiqueta = any(
            nombre_lower.startswith(p) or nombre_lower.endswith(p)
            for p in palabras_clave
        )

        if not ya_tiene_etiqueta:
            # Si NO la tiene, le agregamos el prefijo deseado (ej. "Extracto_12345.xlsx")
            nuevo_nombre = f"{tipo_documento}{nombre_base}{extension_salida}"
        else:
            # Si YA la tiene, mantenemos el nombre base intacto y solo aseguramos la extensión
            nuevo_nombre = f"{nombre_base}{extension_salida}"

        return os.path.join(directorio, nuevo_nombre)

    def process_pdf(self):
        if not self.pdf_engine.has_document:
            messagebox.showwarning("Atención", "Carga un archivo PDF primero.")
            return

        if ejecutar_proceso_exportacion is None:
            messagebox.showerror(
                "Error de Módulo",
                f"No se pudo importar el script de backend:\n{_import_error_msg}",
            )
            return

        nombre_base_pdf = os.path.splitext(
            os.path.basename(self.pdf_engine.current_path)
        )[0]
        nombre_lower = nombre_base_pdf.lower()

        # Determinar el tipo de documento activo
        tipo_actual = getattr(self, "tipo_documento", "Extracto")
        etiqueta_sugerida = f"{tipo_actual}_"

        # Palabras clave ampliadas
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
        self.info_label.config(
            text="⏳ Procesando PDF, por favor espera...", style="Info.TLabel"
        )
        self.root.update_idletasks()

        try:
            excel_generado = ejecutar_proceso_exportacion(
                self.pdf_engine.current_path, output_excel_path=save_path
            )

            self.pdf_engine.last_excel_path = excel_generado

            self.btn_open_excel.config(
                state=tk.NORMAL, style="Secondary.TButton"
            )

            self.info_label.config(
                text=(
                    "✅ Excel"
                    f" generado:\n{os.path.basename(excel_generado)}"
                ),
                style="Status.TLabel",
            )
            ClasificadorConceptos(self.root, excel_generado, self.cargar_excel_en_gui)
            self.cargar_excel_en_gui(excel_generado)

            messagebox.showinfo(
                "Proceso Exitoso",
                "¡Extracción completada!\n\nArchivo guardado"
                f" en:\n{excel_generado}",
            )
        except Exception as e:
            self.info_label.config(
                text="⚠️ Ocurrió un error al procesar el PDF.",
                style="Info.TLabel",
            )
            messagebox.showerror(
                "Error al procesar",
                f"Ocurrió un error en la extracción:\n{str(e)}",
            )
        finally:
            self.root.config(cursor="")

   

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

            # 🟢 Habilitar el botón de limpiar sin llamar a update_idletasks en 'self'
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
        """Limpia la vista previa de las páginas del PDF y resetea el scrollbar

        sin destruir la estructura del visor.
        """
        try:
            # 1. Resetear variables de ruta en el motor y en la GUI
            if hasattr(self, "pdf_engine") and self.pdf_engine:
                if hasattr(self.pdf_engine, "current_path"):
                    self.pdf_engine.current_path = None
                if hasattr(self.pdf_engine, "last_pdf_path"):
                    self.pdf_engine.last_pdf_path = None

            if hasattr(self, "current_pdf_path"):
                self.current_pdf_path = None

            # 2. VACIAR LAS IMÁGENES DENTRO DEL SCROLLABLE_FRAME
            if hasattr(self, "scrollable_frame") and self.scrollable_frame:
                for child in self.scrollable_frame.winfo_children():
                    child.destroy()

            # 3. RESETEAR EL CANVAS Y EL SCROLLBAR
            if hasattr(self, "canvas") and self.canvas:
                # Eliminar cualquier elemento dibujado directamente en el canvas
                self.canvas.delete("all")
                # Volver a vincular la ventana para que el scrollable_frame siga dentro
                if hasattr(self, "scrollable_frame"):
                    self.canvas_window = self.canvas.create_window(
                        (0, 0), window=self.scrollable_frame, anchor="nw"
                    )
                # Resetear la región de scroll a cero
                self.canvas.configure(scrollregion=(0, 0, 0, 0))
                self.canvas.yview_moveto(0)

            # 4. Vaciar referencias de imágenes en memoria para liberar RAM
            if hasattr(self, "pdf_images"):
                self.pdf_images = []
            if hasattr(self, "pdf_image"):
                self.pdf_image = None

            # 5. LIMPIAR PANEL DE ESTADO (Texto del archivo y páginas abajo a la izquierda)
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
        """Libera la vista previa, los dataframes y los recursos asociados al Excel."""
        try:
            # 1. Resetear referencias y DataFrames en el motor
            if hasattr(self, "pdf_engine") and self.pdf_engine:
                if hasattr(self.pdf_engine, "last_excel_path"):
                    self.pdf_engine.last_excel_path = None
                if hasattr(self.pdf_engine, "df_original_raw"):
                    self.pdf_engine.df_original_raw = None

            # 2. Limpiar las pestañas de las hojas cargadas en la GUI (Notebook)
            if hasattr(self, "notebook_hojas") and self.notebook_hojas:
                for tab in self.notebook_hojas.tabs():
                    self.notebook_hojas.forget(tab)

            # 3. Deshabilitar el botón de "Abrir Excel" (si aplica en tu interfaz)
            if hasattr(self, "btn_open_excel") and self.btn_open_excel:
                try:
                    self.btn_open_excel.configure(state="disabled")
                except Exception:
                    self.btn_open_excel.config(state="disabled")

            # 4. Resetear etiqueta de estado del Excel si la utilizas
            if hasattr(self, "lbl_excel_estado") and self.lbl_excel_estado:
                self.lbl_excel_estado.config(text="")

        except Exception as e:
            messagebox.showerror(
                "Error al liberar Excel",
                f"Ocurrió un error al limpiar la vista previa del Excel:\n{e}",
            )
    def limpiar_sesion(self):
        """Limpia la sesión de trabajo activa (PDF, imágenes de vista previa y Excel)."""
        try:
            # Siempre ejecutamos liberar_pdf y liberar_excel
            self.liberar_pdf()
            self.liberar_excel()

            # Deshabilitar el botón nuevamente
            if hasattr(self, "btn_limpiar") and self.btn_limpiar:
                try:
                    self.btn_limpiar.configure(state="disabled")
                except Exception:
                    self.btn_limpiar.config(state="disabled")

            # Deshabilitar botones de acción secundarios (como "Extraer TXT y Excel" o "Abrir Excel")
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