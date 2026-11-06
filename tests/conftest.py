import json
import sys
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parent.parent
if str(RAIZ) not in sys.path:
    sys.path.insert(0, str(RAIZ))


def palabra(x, top, texto, ancho_car=4.0):
    return {"x0": float(x), "x1": float(x) + len(texto) * ancho_car, "top": float(top), "text": texto}


def fila(top, *celdas):
    """celdas: (x, texto); un texto con espacios se divide en palabras consecutivas."""
    salida = []
    for x, texto in celdas:
        for p in texto.split():
            salida.append(palabra(x, top, p))
            x += len(p) * 4.0 + 4.0
    return salida


class PaginaFalsa:
    def __init__(self, filas, numero=1, ancho=612.0, alto=792.0, texto=None):
        self.words = [w for f in filas for w in f]
        self.page_number = numero
        self.width = ancho
        self.height = alto
        self._texto = texto

    def extract_words(self, **_):
        return list(self.words)

    def extract_text(self, **_):
        if self._texto is not None:
            return self._texto
        lineas = {}
        for w in self.words:
            lineas.setdefault(w["top"], []).append(w)
        return "\n".join(" ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"])) for _, ws in sorted(lineas.items()))

    def extract_text_lines(self, **_):
        lineas = {}
        for w in self.words:
            lineas.setdefault(w["top"], []).append(w)
        return [{"top": top, "text": " ".join(w["text"] for w in sorted(ws, key=lambda w: w["x0"]))}
                for top, ws in sorted(lineas.items())]


class PDFFalso:
    def __init__(self, paginas):
        self.pages = paginas

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False


@pytest.fixture
def pdf_falso(monkeypatch):
    """Reemplaza pdfplumber.open por documentos en memoria: pdf_falso(ruta, [PaginaFalsa, ...])."""
    import pdfplumber

    registro = {}
    monkeypatch.setattr(pdfplumber, "open", lambda ruta: PDFFalso(registro[str(ruta)]))

    def registrar(ruta, paginas):
        registro[str(ruta)] = paginas
        return str(ruta)

    return registrar


@pytest.fixture
def ruta_administradoras(tmp_path):
    datos = {"administradoras": [
        {"NOMBRE": "SALUD TOTAL", "NIT": "800130907", "CATEGORIA": "EPS", "CUENTA": "237005001", "CODIGO": "EPS002"},
        {"NOMBRE": "PORVENIR", "NIT": "800224808", "CATEGORIA": "PENSION", "CUENTA": "238030001", "CODIGO": "230301"},
        {"NOMBRE": "PORVENIR CESANTIAS", "NIT": "800224808", "CATEGORIA": "CESANTIAS", "CUENTA": "251005001", "CODIGO": ""},
        {"NOMBRE": "COMFAMA", "NIT": "890900841", "CATEGORIA": "CCF", "CUENTA": "237010001", "CODIGO": "CCF04"},
        {"NOMBRE": "ARL SURA", "NIT": "890903790", "CATEGORIA": "ARL", "CUENTA": "237006001", "CODIGO": "14-11"},
        {"NOMBRE": "SENA", "NIT": "899999034", "CATEGORIA": "PARAFISCALES", "CUENTA": "237010002", "CODIGO": "SENA"},
        {"NOMBRE": "ICBF", "NIT": "899999239", "CATEGORIA": "PARAFISCALES", "CUENTA": "237010003", "CODIGO": "ICBF"},
    ]}
    ruta = tmp_path / "administradoras.json"
    ruta.write_text(json.dumps(datos, ensure_ascii=False), encoding="utf-8")
    return str(ruta)


@pytest.fixture
def sin_conversion_xls(monkeypatch):
    """Sustituye la conversión COM a Excel 5.0/95 por un renombrado (pruebas sin Microsoft Excel)."""
    import os

    import servicios.plano as plano

    monkeypatch.setattr(plano, "convertir_a_xls95", lambda origen, destino: os.replace(origen, destino))