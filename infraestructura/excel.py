import os

import pandas as pd
from openpyxl.styles import Font

from nucleo.constantes import CLAVE_CUENTA, COLUMNAS_MONEDA, FORMATO_MONEDA, HOJA_INFO

NEGRITA = Font(bold=True)
FORMATO_XLS95 = 39


def aplicar_formato_moneda(ws):
    for col in ws.iter_cols(1, ws.max_column):
        if str(col[0].value or "").upper() in COLUMNAS_MONEDA:
            for celda in col[1:]:
                if isinstance(celda.value, (int, float)):
                    celda.number_format = FORMATO_MONEDA


def negrita(ws, filas, columnas):
    for fila in filas:
        for col in columnas:
            ws.cell(fila, col).font = NEGRITA


def anexar_resumenes(ws, bloques, col_concepto, col_monto):
    fila = ws.max_row + 3
    for titulo, resumen in bloques:
        ws.cell(fila, col_concepto, titulo).font = NEGRITA
        for i, (concepto, monto) in enumerate(resumen.itertuples(index=False), start=1):
            ws.cell(fila + i, col_concepto, concepto).font = NEGRITA
            celda = ws.cell(fila + i, col_monto, float(monto))
            celda.number_format = FORMATO_MONEDA
            celda.font = NEGRITA
        fila += len(resumen) + 3


def hoja_info(cuenta):
    return pd.DataFrame({"CLAVE": [CLAVE_CUENTA], "VALOR": [str(cuenta)]})


def leer_cuenta_bancaria(libro):
    """Cuenta guardada en la hoja Info; acepta ruta o pd.ExcelFile."""
    try:
        xl = libro if isinstance(libro, pd.ExcelFile) else pd.ExcelFile(libro)
        if HOJA_INFO not in xl.sheet_names:
            return ""
        info = pd.read_excel(xl, sheet_name=HOJA_INFO, dtype=str).fillna("")
        return next(iter(info.loc[info["CLAVE"] == CLAVE_CUENTA, "VALOR"]), "").strip()
    except Exception:
        return ""


def convertir_a_xls95(ruta_xlsx, ruta_xls):
    """Guarda ruta_xlsx como Excel 5.0/95 mediante Microsoft Excel (COM) y elimina el .xlsx."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError("Falta pywin32: instale con 'pip install pywin32'.") from e
    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(ruta_xlsx))
        wb.SaveAs(os.path.abspath(ruta_xls), FileFormat=FORMATO_XLS95)
        wb.Close(False)
    except Exception as e:
        raise RuntimeError(f"Fallo en la conversión a formato 5.0/95: {e}") from e
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    if os.path.exists(ruta_xlsx):
        os.remove(ruta_xlsx)


def abrir_con_sistema(ruta):
    if ruta and os.path.exists(ruta):
        os.startfile(ruta)
        return True
    return False