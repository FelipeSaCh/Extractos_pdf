import os
import fitz  # PyMuPDF
from PIL import Image, ImageTk

class PDFEngine:
    def __init__(self):
        self.doc = None
        self.current_path = None
        self.last_excel_path = None
        self.page_images = []

    def open_pdf(self, path: str, zoom: float = 1.5):
        """Abre el PDF y prepara las imágenes para Tkinter."""
        self.current_path = path
        self.last_excel_path = None
        self.doc = fitz.open(path)
        self.page_images.clear()

        for page in self.doc:
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            photo = ImageTk.PhotoImage(img)
            self.page_images.append(photo)

        return len(self.doc)

    def open_generated_excel(self):
        """Abre el Excel generado con el programa predeterminado del SO."""
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            os.startfile(self.last_excel_path)
            return True
        return False

    def reorganizar_excel_actual(self, funcion_reorganizar):
        """Ejecuta la función de ordenamiento sobre el Excel actual."""
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            funcion_reorganizar(self.last_excel_path)
            return True
        return False

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None
        self.current_path = None
        self.last_excel_path = None
        self.page_images.clear()

    @property
    def has_document(self) -> bool:
        return self.doc is not None