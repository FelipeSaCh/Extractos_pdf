#pdf_engine.py
import os
import tempfile
import fitz  # PyMuPDF
from PIL import Image


class PDFPasswordRequiredError(Exception):
    """Indica que el PDF está protegido y requiere contraseña."""
    pass


class PDFInvalidPasswordError(Exception):
    """Indica que la contraseña suministrada no es válida."""
    pass


class PDFEngine:
    def __init__(self):
        self.doc = None

        # Ruta que utilizará el resto del programa para procesar el PDF.
        # Si el PDF está protegido, será la ruta del PDF temporal desbloqueado.
        self.current_path = None

        # Ruta original seleccionada por el usuario.
        self.original_path = None

        self.last_pdf_path = None
        self.last_excel_path = None

        # Lista de imágenes PIL (no ImageTk.PhotoImage): la generación es
        # thread-safe y puede ejecutarse en un hilo secundario. La conversión
        # a PhotoImage debe hacerla la GUI en el hilo principal de Tkinter.
        self.page_images = []
        self.df_original_raw = None

        # Ruta del PDF temporal generado cuando el documento
        # original estaba protegido con contraseña.
        self.temp_pdf_path = None

    def open_pdf(self, path: str, password: str = None):
        """
        Abre un PDF y deja el documento listo para renderizar (operación rápida,
        sin generar imágenes de página).

        Comportamiento:
        - PDF normal:
            Se abre directamente.
        - PDF protegido sin contraseña:
            Lanza PDFPasswordRequiredError.
        - PDF protegido con contraseña correcta:
            Genera un PDF temporal sin protección y lo abre.
        - PDF protegido con contraseña incorrecta:
            Lanza PDFInvalidPasswordError.
        """

        if not path:
            raise ValueError("No se proporcionó una ruta de PDF.")

        if not os.path.exists(path):
            raise FileNotFoundError(
                f"No se encontró el archivo PDF:\n{path}"
            )

        # Limpiar cualquier documento anterior.
        self.liberar_recursos()

        self.original_path = path
        self.last_pdf_path = path
        self.last_excel_path = None

        try:
            doc = fitz.open(path)

            # ---------------------------------------------------------
            # PDF PROTEGIDO
            # ---------------------------------------------------------
            if doc.needs_pass:
                # Todavía no se proporcionó contraseña.
                if password is None:
                    doc.close()

                    raise PDFPasswordRequiredError(
                        "El PDF está protegido con contraseña."
                    )

                # Intentar autenticar.
                autenticado = doc.authenticate(password)

                if not autenticado:
                    doc.close()

                    raise PDFInvalidPasswordError(
                        "La contraseña ingresada no es correcta."
                    )

                # -----------------------------------------------------
                # CREAR COPIA TEMPORAL SIN PROTECCIÓN
                # -----------------------------------------------------
                # Extraemos el nombre base sin extensión para conservarlo en el temporal
                nombre_original = os.path.splitext(os.path.basename(path))[0]
                prefix_temp = f"pdfengine_{nombre_original}_"

                temp_file = tempfile.NamedTemporaryFile(
                    prefix=prefix_temp,
                    suffix=".pdf",
                    delete=False
                )

                temp_path = temp_file.name
                temp_file.close()

                try:
                    doc.save(
                        temp_path,
                        encryption=fitz.PDF_ENCRYPT_NONE,
                        garbage=4,
                        deflate=True
                    )

                    doc.close()
                    doc = None

                    self.temp_pdf_path = temp_path

                    # Abrimos nuevamente la copia ya desbloqueada.
                    doc = fitz.open(temp_path)

                except Exception:
                    if doc is not None:
                        doc.close()

                    if os.path.exists(temp_path):
                        try:
                            os.remove(temp_path)
                        except OSError:
                            pass

                    self.temp_pdf_path = None

                    raise

            # ---------------------------------------------------------
            # PDF NORMAL O PDF YA DESBLOQUEADO
            # ---------------------------------------------------------
            self.doc = doc

            # Para PDF normal current_path será el original.
            # Para PDF protegido será el temporal desbloqueado.
            if self.temp_pdf_path:
                self.current_path = self.temp_pdf_path
            else:
                self.current_path = path

            return len(self.doc)

        except (PDFPasswordRequiredError, PDFInvalidPasswordError):
            raise

        except Exception:
            self.liberar_recursos()
            raise

    def renderizar_paginas(self, zoom: float = 1.5, progreso_callback=None):
        """Genera las imágenes PIL de cada página del documento ya abierto.

        Operación pesada (CPU-bound) pero thread-safe: no crea ni toca
        ningún objeto de Tkinter, por lo que puede ejecutarse en un hilo
        secundario. `progreso_callback(indice, total)` se invoca tras
        renderizar cada página, si se proporciona.
        """
        if not self.has_document:
            raise ValueError("No hay un documento PDF abierto.")

        self.page_images.clear()

        matrix = fitz.Matrix(zoom, zoom)
        total = len(self.doc)

        for indice, page in enumerate(self.doc, start=1):
            pix = page.get_pixmap(matrix=matrix)

            img = Image.frombytes(
                "RGB",
                [pix.width, pix.height],
                pix.samples
            )

            self.page_images.append(img)

            if progreso_callback:
                progreso_callback(indice, total)

        return self.page_images

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

    def _eliminar_pdf_temporal(self):
        """Elimina el PDF temporal desbloqueado si existe."""

        if self.temp_pdf_path:
            try:
                if os.path.exists(self.temp_pdf_path):
                    os.remove(self.temp_pdf_path)
            except OSError:
                # Si Windows todavía tiene el archivo bloqueado,
                # no detenemos el programa.
                pass

            self.temp_pdf_path = None

    def liberar_recursos(self):
        """
        Reset total de variables para que el motor olvide
        los archivos cargados y elimine cualquier PDF temporal.
        """

        # Primero cerrar el documento para liberar el bloqueo
        # sobre el archivo temporal.
        if self.doc:
            try:
                self.doc.close()
            except Exception:
                pass

            self.doc = None

        # Después de cerrar PyMuPDF podemos eliminar el temporal.
        self._eliminar_pdf_temporal()

        self.current_path = None
        self.original_path = None
        self.last_pdf_path = None
        self.last_excel_path = None

        self.page_images.clear()
        self.df_original_raw = None

    def close(self):
        """Alias de liberar_recursos()."""
        self.liberar_recursos()

    @property
    def has_document(self) -> bool:
        return self.doc is not None

    @property
    def is_temporary_document(self) -> bool:
        """Indica si el PDF actualmente abierto es una copia temporal."""
        return self.temp_pdf_path is not None