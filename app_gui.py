import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from pdf_engine import PDFEngine
from version import __version__

try:
    from script import ejecutar_proceso_exportacion, reorganizar_excel
except ImportError as err:
    ejecutar_proceso_exportacion = None
    reorganizar_excel = None
    _import_error_msg = str(err)


class PDFViewerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"Lector y Extractor Bancolombia v{__version__} - Tkinter")
        self.root.geometry("1050x700")

        self.pdf_engine = PDFEngine()
        self._build_ui()

    def _build_ui(self):
        paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        paned_window.pack(fill=tk.BOTH, expand=True)

        # PANEL IZQUIERDO: Controles
        left_frame = ttk.Frame(paned_window, width=260, padding=15)
        paned_window.add(left_frame, weight=1)

        title_label = ttk.Label(
            left_frame, 
            text="Panel de Control", 
            font=("Helvetica", 14, "bold")
        )
        title_label.pack(anchor="w", pady=(0, 15))

        btn_load = ttk.Button(
            left_frame, 
            text="📁 Cargar PDF", 
            command=self.load_pdf
        )
        btn_load.pack(fill=tk.X, pady=5)

        self.btn_process = ttk.Button(
            left_frame, 
            text="⚡ Extraer TXT y Excel", 
            command=self.process_pdf,
            state=tk.DISABLED
        )
        self.btn_process.pack(fill=tk.X, pady=5)

        self.btn_reorganize = ttk.Button(
            left_frame, 
            text="🔄 Reorganizar Excel", 
            command=self.reorganize_excel_file,
            state=tk.DISABLED
        )
        self.btn_reorganize.pack(fill=tk.X, pady=5)

        self.btn_open_excel = ttk.Button(
            left_frame, 
            text="🟢 Abrir Excel", 
            command=self.open_excel,
            state=tk.DISABLED
        )
        self.btn_open_excel.pack(fill=tk.X, pady=5)

        self.info_label = ttk.Label(
            left_frame, 
            text="Ningún archivo cargado", 
            wraplength=220, 
            foreground="gray"
        )
        self.info_label.pack(anchor="w", pady=15)

        # PANEL DERECHO: Visualizador
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=4)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.tab_pdf = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_pdf, text="  Vista Previa PDF  ")
        

        self.canvas = tk.Canvas(self.tab_pdf, bg="#525659")
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

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def load_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Seleccionar extracto PDF",
            filetypes=[("Archivos PDF", "*.pdf")]
        )

        if not file_path:
            return

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
                foreground="black"
            )
            self.btn_process.config(state=tk.NORMAL)
            self.btn_reorganize.config(state=tk.DISABLED)
            self.btn_open_excel.config(state=tk.DISABLED)

        except Exception as e:
            messagebox.showerror("Error de Carga", f"No se pudo cargar el PDF:\n{str(e)}")

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

        try:
            # Pasa la ruta seleccionada
            excel_generado = ejecutar_proceso_exportacion(self.pdf_engine.current_path, output_excel_path=save_path)
            
            self.pdf_engine.last_excel_path = excel_generado

            self.btn_reorganize.config(state=tk.NORMAL)
            self.btn_open_excel.config(state=tk.NORMAL)

            messagebox.showinfo(
                "Proceso Exitoso", 
                f"¡Extracción completada!\n\nArchivo guardado en:\n{excel_generado}"
            )
        except Exception as e:
            messagebox.showerror("Error al procesar", f"Ocurrió un error en la extracción:\n{str(e)}")

    def reorganize_excel_file(self):
        if reorganizar_excel is None:
            messagebox.showerror("Error de Módulo", "La función 'reorganizar_excel' no está disponible.")
            return

        try:
            exito = self.pdf_engine.reorganizar_excel_actual(reorganizar_excel)
            if exito:
                messagebox.showinfo(
                    "Reorganización Exitosa", 
                    "Se han movido las filas con valores negativos a la parte superior del Excel."
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

    def open_excel(self):
        if not self.pdf_engine.open_generated_excel():
            messagebox.showerror("Error", "No se encontró el archivo Excel.")

    def _on_mousewheel(self, event):
        if self.pdf_engine.has_document:
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFViewerApp(root)
    root.mainloop()