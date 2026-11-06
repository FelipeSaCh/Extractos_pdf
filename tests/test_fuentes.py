import os
from types import SimpleNamespace

import pytest

from extractores import Extraccion
from servicios.fuentes import (BTN_EXTRACTO, BTN_PLANILLA, FUENTES, Fuente, con_prefijo,
                               procesar_excel_bancolombia)


def test_registro_modulos_cargan():
    for fuente in FUENTES.values():
        modulo = fuente.modulo()
        assert hasattr(modulo, "Generador" if fuente.planilla else "extraer"), fuente.nombre


def test_botones_por_fuente():
    assert FUENTES["Bancolombia (Estándar)"].botones[0] == BTN_EXTRACTO
    assert all(f.botones == (BTN_PLANILLA,) for f in FUENTES.values() if f.planilla)
    assert FUENTES["Seguridad social (Aportes en Línea)"].acepta_excel


def test_opciones_davivienda():
    assert "tipo_cuenta" in FUENTES["Davivienda"].opciones()
    assert FUENTES["Nequi"].opciones() == {}


def test_error_de_carga_identifica_la_fuente():
    def falla():
        raise ModuleNotFoundError("falta x")
    with pytest.raises(ImportError, match="Banco Prueba"):
        Fuente("Banco Prueba", falla, ()).modulo()


def test_exportar_sin_movimientos():
    modulo = SimpleNamespace(extraer=lambda *a, **k: Extraccion(["A"], []), SIN_MOVIMIENTOS="Nada que leer")
    with pytest.raises(ValueError, match="Nada que leer"):
        Fuente("X", lambda: modulo, ()).exportar("x.pdf")


def test_exportar_pasa_opciones(tmp_path):
    recibido = {}

    def extraer(ruta, progreso, **op):
        recibido.update(op)
        return Extraccion(["FECHA", "DESCRIPCIÓN", "VALOR"], ["1/05|||A|||10"])

    ruta = Fuente("X", lambda: SimpleNamespace(extraer=extraer), ()).exportar(
        str(tmp_path / "e.pdf"), str(tmp_path / "e.xlsx"), tipo_cuenta="DAMAS")
    assert recibido == {"tipo_cuenta": "DAMAS"} and os.path.exists(ruta)


def test_con_prefijo(tmp_path):
    original = tmp_path / "abril.pdf"
    original.write_bytes(b"")
    nueva = con_prefijo(str(original), "Extracto")
    assert os.path.basename(nueva) == "Extracto_abril.pdf" and not original.exists()
    assert con_prefijo(nueva, "Extracto") == nueva


def test_procesar_excel_bancolombia(tmp_path):
    import openpyxl
    import pandas as pd

    wb = openpyxl.Workbook()
    for f in [["SALDO ANTERIOR", None, "SALDO ACTUAL"], [1000, None, 1170],
              ["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO", "VALOR", "SALDO"],
              ["2/05", "PAGO", "MED", "900123456", -30, 970], ["8/05", "ABONO", "", "", "200,00", 1170]]:
        wb.active.append(f)
    entrada = tmp_path / "banco.xlsx"
    wb.save(entrada)
    salida = procesar_excel_bancolombia(str(entrada))
    assert salida.endswith("banco_procesado.xlsx")
    assert "DCTO." in pd.read_excel(salida, sheet_name="Datos").columns
    resumen = dict(pd.read_excel(salida, sheet_name="Resumen").itertuples(index=False))
    assert resumen["DIFERENCIA CONCILIACIÓN"] == 0