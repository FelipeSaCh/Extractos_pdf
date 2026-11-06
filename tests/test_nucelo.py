import math

import pandas as pd
import pytest

from nucleo.fechas import mes_de_fechas, partes_fecha
from nucleo.numeros import a_entero, a_float, serie_a_numero
from nucleo.texto import clave, codigo, digitos, normalizar, sin_tildes
from nucleo.validaciones import (PATRON_COMPROBANTE, PATRON_CUENTA, PATRON_NIT, PATRON_TERCERO,
                                 PATRON_TERCERO_REFERENCIA)


@pytest.mark.parametrize("texto, esperado", [
    ("1.234.567,89", 1234567.89),
    ("1,234,567.89", 1234567.89),
    ("-$1,234.56", -1234.56),
    ("$ -50.000", -50000.0),
    ("-$50.000", -50000.0),
    ("50.000", 50000.0),
    ("1,234.56-", -1234.56),
    ("(1.200)", -1200.0),
    ("$ 1.234.567", 1234567.0),
    ("123.45", 123.45),
    ("1.5", 1.5),
    ("0,00", 0.0),
    (1500, 1500.0),
    (-12.5, -12.5),
])
def test_a_float_formatos(texto, esperado):
    assert a_float(texto) == pytest.approx(esperado)


@pytest.mark.parametrize("vacio", [None, "", "abc", "$", "-", float("nan")])
def test_a_float_valor_por_defecto(vacio):
    assert a_float(vacio) == 0.0
    assert math.isnan(a_float(vacio, math.nan))


def test_serie_a_numero_texto_y_numerica():
    texto = serie_a_numero(pd.Series(["1,000.50", "-2.000", "x"]))
    assert texto.iloc[0] == 1000.5 and texto.iloc[1] == -2000.0 and math.isnan(texto.iloc[2])
    assert serie_a_numero(pd.Series([1, 2.5])).tolist() == [1.0, 2.5]


@pytest.mark.parametrize("valor, esperado", [("7", 7), (" 12 ", 12), ("0", 1), ("-3", 1), ("abc", 1), (None, 1)])
def test_a_entero(valor, esperado):
    assert a_entero(valor) == esperado


@pytest.mark.parametrize("valor, esperado", [
    ("2026/05/03", (2026, 5, 3)),
    ("2026-05-03 00:00:00", (2026, 5, 3)),
    ("05/03/26", (2026, 3, 5)),
    ("05/03/2026", (2026, 3, 5)),
    ("5 ene 2026", (2026, 1, 5)),
    (pd.Timestamp("2026-07-09"), (2026, 7, 9)),
])
def test_partes_fecha_con_ano(valor, esperado):
    assert partes_fecha(valor) == esperado


@pytest.mark.parametrize("valor, mes, dia", [("5/03", 3, 5), ("12 dic.", 12, 12), (5.03, 3, 5)])
def test_partes_fecha_sin_ano(valor, mes, dia):
    ano, m, d = partes_fecha(valor)
    assert math.isnan(ano) and (m, d) == (mes, dia)


@pytest.mark.parametrize("valor", [None, "", "sin fecha", float("nan")])
def test_partes_fecha_invalida(valor):
    assert all(math.isnan(p) for p in partes_fecha(valor))


def test_mes_de_fechas():
    assert mes_de_fechas(pd.Series(["1/05", "2026/06/01", "3 jul 2026", None, "99/99"])).tolist() == [5, 6, 7, 0, 0]


def test_texto():
    assert sin_tildes("Planilla Única") == "PLANILLA UNICA"
    assert normalizar("  Caja de Compensación-Comfama ") == "CAJA DE COMPENSACION COMFAMA"
    assert clave("Cód. Pensión") == "CODPENSION"
    assert codigo(" eps 002 ") == "EPS002"
    assert digitos("900.123.456-7") == "9001234567"
    assert digitos(900123456.0) == "900123456"
    assert digitos(float("nan")) == ""


def test_validaciones():
    assert PATRON_COMPROBANTE.match("EG1") and not PATRON_COMPROBANTE.match("EGRESO")
    assert PATRON_CUENTA.match("112005001") and not PATRON_CUENTA.match("11A")
    assert PATRON_TERCERO.match("900123456-7") and PATRON_TERCERO.match("12345")
    assert not PATRON_TERCERO_REFERENCIA.match("12345")
    assert not PATRON_NIT.match("900123456-7")