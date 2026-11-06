"""Re-estructuración del extracto Bancolombia descargado en Excel."""
import re

import openpyxl

from extractores import Extraccion
from nucleo.constantes import DELIMITADOR
from nucleo.numeros import a_float

COLUMNAS = ["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO.", "VALOR", "SALDO"]
_FECHA_CORTA = re.compile(r"^\d{1,2}[/-]\d{1,2}$")
_ETIQUETAS_SALDO = {"SALDO ANTERIOR": 0, "SALDO ACTUAL": 1}


def _saldos(data):
    """Valor ubicado debajo de los encabezados SALDO ANTERIOR / SALDO ACTUAL."""
    saldos = [None, None]
    for r, fila in enumerate(data[:-1]):
        for c, celda in enumerate(fila):
            texto = str(celda or "").strip().upper()
            for etiqueta, i in _ETIQUETAS_SALDO.items():
                if etiqueta in texto and c < len(data[r + 1]):
                    saldos[i] = a_float(data[r + 1][c])
    return saldos


def _celda(fila, i):
    return str(fila[i]).strip() if len(fila) > i and fila[i] else ""


def extraer(excel_path, progreso=None, **_):
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        data = list(wb.active.iter_rows(values_only=True))
    finally:
        wb.close()
    if not data:
        raise ValueError("El archivo Excel está vacío.")

    lineas, en_movimientos = [], False
    for fila in data:
        if not fila or not any(c is not None and str(c).strip() for c in fila):
            continue
        linea = " ".join(str(c).upper() for c in fila if c is not None)
        if "FECHA" in linea and "DESCRIPCI" in linea:
            en_movimientos = True
            continue
        fecha = _celda(fila, 0)
        if not en_movimientos or not _FECHA_CORTA.match(fecha):
            continue
        valor = a_float(fila[4]) if len(fila) > 4 else 0.0
        saldo = a_float(fila[5]) if len(fila) > 5 else 0.0
        lineas.append(DELIMITADOR.join([fecha, _celda(fila, 1), _celda(fila, 2), _celda(fila, 3),
                                        f"{valor:.2f}", f"{saldo:.2f}"]))
    return Extraccion(COLUMNAS, lineas, "", *_saldos(data))