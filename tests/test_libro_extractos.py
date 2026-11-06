import openpyxl
import pandas as pd
import pytest

from extractores import Extraccion
from infraestructura.excel import leer_cuenta_bancaria
from servicios.libro_extracto import escribir_libro, exportar

COLUMNAS = ["FECHA", "DESCRIPCIÓN", "VALOR", "SALDO"]


def _df(filas):
    return pd.DataFrame(filas, columns=COLUMNAS)


def _resumen(ruta):
    return dict(pd.read_excel(ruta, sheet_name="Resumen").itertuples(index=False))


def test_un_mes_cargos_primero_y_resumen(tmp_path):
    ruta = escribir_libro(_df([["1/05", "ABONO", "100.00", "1,100.00"], ["2/05", "PAGO", "-30.00", "1,070.00"]]),
                          tmp_path / "a.xlsx", cuenta_bancaria="123456789")
    datos = pd.read_excel(ruta, sheet_name="Datos")
    assert datos["DESCRIPCIÓN"].tolist()[:2] == ["PAGO", "ABONO"]
    assert _resumen(ruta) == {"SALDO ANTERIOR": 1000, "TOTAL CARGOS": -30, "TOTAL ABONOS": 100, "SALDO ACTUAL": 1070}
    assert leer_cuenta_bancaria(ruta) == "123456789"
    assert openpyxl.load_workbook(ruta).sheetnames == ["Datos", "Resumen", "Conceptos", "Info"]


def test_conciliacion_con_saldos_del_banco(tmp_path):
    df = pd.DataFrame([["1/05", "ABONO", "200"], ["2/05", "PAGO", "-60"]], columns=["FECHA", "DESCRIPCIÓN", "VALOR"])
    ruta = escribir_libro(df, tmp_path / "b.xlsx", saldo_anterior=1000, saldo_actual=1150)
    assert _resumen(ruta)["DIFERENCIA CONCILIACIÓN"] == pytest.approx(-10)


def test_varios_meses_bloques_por_mes(tmp_path):
    ruta = escribir_libro(_df([["1/05", "A", "10", "110"], ["1/06", "B", "-5", "105"]]), tmp_path / "c.xlsx")
    desc = pd.read_excel(ruta, sheet_name="Datos")["DESCRIPCIÓN"].dropna().tolist()
    assert "RESUMEN MAYO" in desc and "RESUMEN JUNIO" in desc and desc[-1] != "A"
    assert "RESUMEN GENERAL" in desc
    conceptos = pd.read_excel(ruta, sheet_name="Conceptos")
    total = conceptos[conceptos["DESCRIPCIÓN"] == "TOTAL GENERAL"].iloc[0]
    assert (total["CARGOS"], total["ABONOS"], total["NETO"]) == (-5, 10, 5)


def test_exportar_genera_txt_y_xlsx(tmp_path):
    e = Extraccion(COLUMNAS, ["1/05|||ABONO|||100.00|||100.00"], "999888777")
    ruta = exportar(e, str(tmp_path / "Extracto.pdf"))
    assert ruta.endswith("Extracto_convertido.xlsx")
    assert (tmp_path / "Extracto_convertido_plano.txt").read_text(encoding="utf-8").startswith("FECHA|||")
    assert exportar(Extraccion(COLUMNAS, []), str(tmp_path / "vacio.pdf")) is None