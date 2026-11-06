import os
import tempfile

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c


class PDFPasswordRequiredError(Exception):
    """El PDF está protegido y requiere contraseña."""


class PDFInvalidPasswordError(Exception):
    """La contraseña suministrada no es válida."""


def _es_error_password(error):
    return getattr(error, "err_code", None) == pdfium_c.FPDF_ERR_PASSWORD


class PDFEngine:
    """Apertura (con desbloqueo a copia temporal) y renderizado de PDF.

    current_path es la ruta que debe usar el resto del programa: la original o la copia desbloqueada.
    page_images contiene imágenes PIL (thread-safe); la conversión a PhotoImage la hace la GUI.
    """

    def __init__(self):
        self.doc = None
        self.current_path = None
        self.original_path = None
        self.temp_pdf_path = None
        self.page_images = []

    def open_pdf(self, path: str, password: str = None):
        if not path:
            raise ValueError("No se proporcionó una ruta de PDF.")
        if not os.path.exists(path):
            raise FileNotFoundError(f"No se encontró el archivo PDF:\n{path}")

        self.liberar_recursos()
        self.original_path = path
        try:
            self.doc = self._abrir(path, password)
            self.current_path = self.temp_pdf_path or path
            return len(self.doc)
        except (PDFPasswordRequiredError, PDFInvalidPasswordError):
            raise
        except Exception:
            self.liberar_recursos()
            raise

    def _abrir(self, path, password):
        try:
            return pdfium.PdfDocument(path)
        except pdfium.PdfiumError as e:
            if not _es_error_password(e):
                raise
        if password is None:
            raise PDFPasswordRequiredError("El PDF está protegido con contraseña.")
        try:
            doc = pdfium.PdfDocument(path, password=password)
        except pdfium.PdfiumError as e:
            if _es_error_password(e):
                raise PDFInvalidPasswordError("La contraseña ingresada no es correcta.") from None
            raise
        return self._desbloquear(doc, path)

    def _desbloquear(self, doc, path):
        nombre = os.path.splitext(os.path.basename(path))[0]
        fd, temp_path = tempfile.mkstemp(prefix=f"pdfengine_{nombre}_", suffix=".pdf")
        os.close(fd)
        try:
            doc.save(temp_path, flags=pdfium_c.FPDF_REMOVE_SECURITY)
            doc.close()
            self.temp_pdf_path = temp_path
            return pdfium.PdfDocument(temp_path)
        except Exception:
            try:
                doc.close()
            except Exception:
                pass
            self.temp_pdf_path = None
            if os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass
            raise

    def renderizar_paginas(self, zoom: float = 1.5, progreso_callback=None):
        if not self.has_document:
            raise ValueError("No hay un documento PDF abierto.")
        self.page_images.clear()
        total = len(self.doc)
        for indice in range(total):
            page = self.doc[indice]
            try:
                self.page_images.append(page.render(scale=zoom).to_pil().convert("RGB"))
            finally:
                page.close()
            if progreso_callback:
                progreso_callback(indice + 1, total)
        return self.page_images

    def _eliminar_pdf_temporal(self):
        if self.temp_pdf_path:
            try:
                if os.path.exists(self.temp_pdf_path):
                    os.remove(self.temp_pdf_path)
            except OSError:
                pass
            self.temp_pdf_path = None

    def liberar_recursos(self):
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass
            self.doc = None
        self._eliminar_pdf_temporal()
        self.current_path = None
        self.original_path = None
        self.page_images.clear()

    close = liberar_recursos

    @property
    def has_document(self) -> bool:
        return self.doc is not None

    @property
    def is_temporary_document(self) -> bool:
        return self.temp_pdf_path is not None