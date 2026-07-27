import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from pdf_engine import PDFEngine
from version import __version__

try:
    from script import ejecutar_proceso_exportacion, reorganizar_excel
except ImportError as err:
    ejecutar_proceso_exportacion = None
    reorganizar_excel = None
    _import_error_msg = str(err)


# ------------------------------------------------------------------
# Paleta y constantes visuales (solo presentación, no afecta lógica)
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

    # ------------------------------------------------------------------
    # Utilidades visuales
    # ------------------------------------------------------------------
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

        # Botón primario (acción principal de carga)
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
            background=[("active", COLOR_PRIMARY_DARK), ("disabled", "#B9C6E4")],
            foreground=[("disabled", "#F0F0F0")],
        )

        # Botón secundario (acciones dependientes)
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
        style.configure(
            "TNotebook",
            background=COLOR_BG,
            borderwidth=0,
        )
        style.configure(
            "TNotebook.Tab",
            font=FONT_BUTTON,
            padding=(16, 8),
        )
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
        style.map("Treeview", background=[("selected", COLOR_PRIMARY)], foreground=[("selected", "white")])

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
                "⚠️ Recuerde que este programa únicamente puede usarse con archivos "
                "PDF que contengan texto (no imágenes escaneadas), y que el soporte "
                "actual es exclusivo para extractos con el formato de Bancolombia. "
                "No se garantiza compatibilidad con otros bancos u otros formatos. "
                "En futuras actualizaciones se trabajará en su estructuración."
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

        # ============= PANEL IZQUIERDO: Controles =============
        left_frame = ttk.Frame(paned_window, width=270, padding=20, style="Sidebar.TFrame")
        left_frame.pack_propagate(False)
        paned_window.add(left_frame, weight=1)

        title_label = ttk.Label(
            left_frame,
            text="📊 Panel de Control",
            style="Title.TLabel",
        )
        title_label.pack(anchor="w", pady=(0, 2))

        subtitle_label = ttk.Label(
            left_frame,
            text="Extrae y organiza tus extractos",
            style="Subtitle.TLabel",
        )
        subtitle_label.pack(anchor="w", pady=(0, 18))

        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=(0, 18))

        self._build_warning_box(left_frame)

        btn_load = ttk.Button(
            left_frame,
            text="📄 Cargar PDF de extractos",
            command=lambda: self.cargar_y_procesar_pdf("Extracto"),
            style="Primary.TButton",
            cursor="hand2",
        )
        btn_load.pack(fill=tk.X, pady=6)

        btn_load_movimientos = ttk.Button(
            left_frame,
            text="📄 Cargar PDF de movimientos",
            command=lambda: self.cargar_y_procesar_pdf("Movimientos"),
            style="Primary.TButton",
            cursor="hand2",
        )
        btn_load_movimientos.pack(fill=tk.X, pady=6)
        

        self.btn_process = ttk.Button(
            left_frame,
            text="⚡  Extraer TXT y Excel",
            command=self.process_pdf,
            state=tk.DISABLED,
            style="Secondary.TButton",
            cursor="hand2",
        )
        self.btn_process.pack(fill=tk.X, pady=6)

        self.btn_reorganize = ttk.Button(
            left_frame,
            text="🔄  Reorganizar Excel",
            command=self.reorganize_excel_file,
            state=tk.DISABLED,
            style="Secondary.TButton",
            cursor="hand2",
        )
        self.btn_reorganize.pack(fill=tk.X, pady=6)

        self.btn_open_excel = ttk.Button(
            left_frame,
            text="🟢  Abrir Excel",
            command=self.open_excel,
            state=tk.DISABLED,
            style="Secondary.TButton",
            cursor="hand2",
        )
        self.btn_open_excel.pack(fill=tk.X, pady=6)

        ttk.Separator(left_frame, orient="horizontal").pack(fill=tk.X, pady=18)

        status_header = ttk.Label(left_frame, text="ESTADO", style="Subtitle.TLabel")
        status_header.pack(anchor="w", pady=(0, 6))

        self.info_label = ttk.Label(
            left_frame,
            text="Ningún archivo cargado",
            wraplength=220,
            style="Info.TLabel",
            justify="left",
        )
        self.info_label.pack(anchor="w", fill=tk.X)

        # Empuja el footer hacia abajo
        spacer = ttk.Frame(left_frame, style="Sidebar.TFrame")
        spacer.pack(fill=tk.BOTH, expand=True)

        footer_label = ttk.Label(
            left_frame,
            text=f"Lector de extractos • v{__version__}",
            style="Footer.TLabel",
        )
        footer_label.pack(anchor="w", side=tk.BOTTOM, pady=(10, 0))

        # ============= PANEL DERECHO: Visualizador =============
        right_frame = ttk.Frame(paned_window, style="Content.TFrame")
        paned_window.add(right_frame, weight=4)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_pdf = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pdf, text="  📄 Vista Previa PDF  ")

        self.canvas = tk.Canvas(self.tab_pdf, bg=COLOR_CANVAS_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(
            self.tab_pdf,
            orient=tk.VERTICAL,
            command=self.canvas.yview
        )

        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.tab_excel = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_excel, text="  📈 Excel Generado  ")

        tree_container = ttk.Frame(self.tab_excel, style="Content.TFrame")
        tree_container.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        self.tree_excel = ttk.Treeview(
            tree_container,
            columns=("FECHA", "DESCRIPCION", "SUCURSAL", "DCTO.", "VALOR", "SALDO"),
            show="headings",
        )
        scrollbar_excel = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.tree_excel.yview)
        self.tree_excel.configure(yscrollcommand=scrollbar_excel.set)
        self.tree_excel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_excel.pack(side=tk.RIGHT, fill=tk.Y)

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        scrollbar_x = ttk.Scrollbar(self.tab_excel, orient=tk.HORIZONTAL, command=self.tree_excel.xview)
        self.tree_excel.configure(xscrollcommand=scrollbar_x.set)
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)

        # Filas alternadas (zebra) para mejorar la lectura de la tabla
        self.tree_excel.tag_configure("odd", background="#F7F9FC")
        self.tree_excel.tag_configure("even", background="white")

    # ------------------------------------------------------------------
    # Lógica de la aplicación (sin cambios funcionales)
    # ------------------------------------------------------------------
    def cargar_y_procesar_pdf(self, prefijo_tipo):
        """Abre el diálogo de archivos, renombra el PDF con el prefijo según el botón

        presionado y llama al cargador de PDF.
        """
        file_path = filedialog.askopenfilename(
            title=f"Seleccionar PDF de {prefijo_tipo.lower()}",
            filetypes=[("Archivos PDF", "*.pdf")],
        )

        if not file_path:
            return

        try:
            folder_path, old_filename = os.path.split(file_path)

            # Asigna el nuevo nombre (Ej: "Extracto_NombreOriginal.pdf")
            nuevo_nombre = f"{prefijo_tipo}_{old_filename}"
            new_file_path = os.path.join(folder_path, nuevo_nombre)

            # Renombra el archivo en el sistema operativo si no tiene ya el prefijo
            if file_path != new_file_path:
                os.replace(file_path, new_file_path)
                file_path = new_file_path

            # Guardar la ruta activa y el tipo en la clase para su posterior uso en reorganizar_excel
            self.current_pdf_path = file_path
            self.tipo_documento = prefijo_tipo

            # Llamar a la función de carga pasándole la nueva ruta
            self.load_pdf(file_path)

        except Exception as e:
            messagebox.showerror(
                "Error al renombrar",
                f"No se pudo asignar el nombre al archivo:\n{str(e)}",
            )


    def load_pdf(self, file_path):
        """Carga y muestra el contenido del PDF en la interfaz gráfica."""
        try:
            total_pages = self.pdf_engine.open_pdf(file_path)

            for child in self.scrollable_frame.winfo_children():
                child.destroy()

            for photo in self.pdf_engine.page_images:
                lbl_page = ttk.Label(self.scrollable_frame, image=photo)
                lbl_page.pack(pady=10, padx=20)

            filename = os.path.basename(file_path)
            self.info_label.config(
                text=f"📄 {filename}\nTotal páginas: {total_pages}",
                style="Status.TLabel",
            )
            self.btn_process.config(state=tk.NORMAL, style="Primary.TButton")
            self.btn_reorganize.config(state=tk.DISABLED, style="Secondary.TButton")
            self.btn_open_excel.config(
                state=tk.DISABLED, style="Secondary.TButton"
            )

        except Exception as e:
            messagebox.showerror(
                "Error de Carga", f"No se pudo cargar el PDF:\n{str(e)}"
            )
    def process_pdf(self):
        if not self.pdf_engine.has_document:
            messagebox.showwarning("Atención", "Carga un archivo PDF primero.")
            return

        if ejecutar_proceso_exportacion is None:
            messagebox.showerror(
                "Error de Módulo",
                f"No se pudo importar 'script.py':\n{_import_error_msg}"
            )
            return

        # Nombre por defecto sugerido
        nombre_sugerido = os.path.splitext(os.path.basename(self.pdf_engine.current_path))[0] + "_convertido.xlsx"

        # Ventana modal para solicitar la ruta y el nombre del archivo de salida
        save_path = filedialog.asksaveasfilename(
            title="Guardar archivo Excel como...",
            initialfile=nombre_sugerido,
            defaultextension=".xlsx",
            filetypes=[("Libro de Excel", "*.xlsx")]
        )

        # Si el usuario cancela la ventana de guardado
        if not save_path:
            return

        # Feedback visual mientras se procesa (no cambia el flujo, solo la UX)
        self.root.config(cursor="watch")
        self.info_label.config(text="⏳ Procesando PDF, por favor espera...", style="Info.TLabel")
        self.root.update_idletasks()

        try:
            # Pasa la ruta seleccionada
            excel_generado = ejecutar_proceso_exportacion(self.pdf_engine.current_path, output_excel_path=save_path)

            self.pdf_engine.last_excel_path = excel_generado

            self.btn_reorganize.config(state=tk.NORMAL, style="Primary.TButton")
            self.btn_open_excel.config(state=tk.NORMAL, style="Secondary.TButton")

            self.info_label.config(
                text=f"✅ Excel generado:\n{os.path.basename(excel_generado)}",
                style="Status.TLabel",
            )

            messagebox.showinfo(
                "Proceso Exitoso",
                f"¡Extracción completada!\n\nArchivo guardado en:\n{excel_generado}"
            )
        except Exception as e:
            self.info_label.config(text="⚠️ Ocurrió un error al procesar el PDF.", style="Info.TLabel")
            messagebox.showerror("Error al procesar", f"Ocurrió un error en la extracción:\n{str(e)}")
        finally:
            self.root.config(cursor="")

    def reorganize_excel_file(self):
        if reorganizar_excel is None:
            messagebox.showerror("Error de Módulo", "La función 'reorganizar_excel' no está disponible.")
            return

        try:
            exito = self.pdf_engine.reorganizar_excel_actual(reorganizar_excel)
            if exito:
                self._cargar_excel_en_treeview(self.pdf_engine.last_excel_path)
                messagebox.showinfo(
                    "Reorganización Exitosa",
                    "El archivo Excel se reorganizó correctamente."
                )
            else:
                messagebox.showwarning("Atención", "No se encontró el Excel generado para reorganizar.")
        except PermissionError:
            messagebox.showerror(
                "Error de Permiso",
                "El archivo Excel está abierto en otro programa. Ciérralo e intenta de nuevo."
            )
        except Exception as e:
            messagebox.showerror("Error al reorganizar", f"Ocurrió un fallo:\n{str(e)}")

    def _cargar_excel_en_treeview(self, file_path):
        if not file_path or not os.path.exists(file_path):
            return

        try:
            # 1. Leer el archivo con pandas
            df = pd.read_excel(file_path)
            df = df.fillna("")  # Reemplazar valores vacíos para evitar errores visuales

            # 2. Limpiar filas antiguas que existan en la tabla
            for item in self.tree_excel.get_children():
                self.tree_excel.delete(item)

            # 3. Ajustar las columnas si el Excel trae columnas dinámicas
            cols = list(df.columns)
            self.tree_excel["columns"] = cols
            self.tree_excel["show"] = "headings"

            for col in cols:
                self.tree_excel.heading(col, text=col)
                self.tree_excel.column(col, width=110, anchor="center")

            # 4. Insertar los datos (con filas alternadas para mejor lectura)
            for i, (_, row) in enumerate(df.iterrows()):
                tag = "odd" if i % 2 else "even"
                self.tree_excel.insert("", tk.END, values=list(row), tags=(tag,))

            # 5. Enfocar/Seleccionar la pestaña de Excel automáticamente
            self.notebook.select(self.tab_excel)

        except Exception as e:
            messagebox.showwarning("Aviso", f"No se pudo cargar la vista previa del Excel:\n{e}")

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
        pass  # Si el ícono no existe o no es válido, se continúa sin él

    app = PDFViewerApp(root)
    root.mainloop()