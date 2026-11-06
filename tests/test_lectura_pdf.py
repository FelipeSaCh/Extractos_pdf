from infraestructura.lectura_pdf import (agrupar_filas, buscar_cuenta_bancaria, es_numero_financiero,
                                         leer_paginas, limpiar_pie_pagina)
from tests.conftest import PaginaFalsa, fila, palabra


def test_buscar_cuenta_bancaria():
    assert buscar_cuenta_bancaria("CUENTA DE AHORROS No. 123-456789-01") == "123-456789-01"
    assert buscar_cuenta_bancaria("Número de cuenta\n3001234567") == "3001234567"
    assert buscar_cuenta_bancaria("CUENTA No 12345") == ""
    assert buscar_cuenta_bancaria(None) == ""


def test_es_numero_financiero():
    assert es_numero_financiero("-1,500.00") and es_numero_financiero("$ 2.000")
    assert not es_numero_financiero("PAGO") and not es_numero_financiero("")


def test_limpiar_pie_pagina():
    filas = fila(700, (40, "TOTAL"), (300, "Página 1 de 3"))
    assert [w["text"] for w in limpiar_pie_pagina(filas)] == ["TOTAL"]
    assert limpiar_pie_pagina(fila(700, (40, "DEFENSOR DEL CONSUMIDOR")), ("DEFENSOR",)) == []


def test_agrupar_filas_por_altura_y_orden_x():
    words = [palabra(200, 100.4, "B"), palabra(40, 99.8, "A"), palabra(40, 120, "C")]
    assert [[w["text"] for w in f] for f in agrupar_filas(words, 3.0)] == [["A", "B"], ["C"]]


def test_leer_paginas_progreso_y_cuenta(pdf_falso):
    paginas = [PaginaFalsa([fila(50, (40, "CUENTA CORRIENTE 000123456789"))], 1), PaginaFalsa([], 2)]
    ruta = pdf_falso("x.pdf", paginas)
    avances = []
    lineas, cuenta = leer_paginas(ruta, lambda p: [f"p{p.page_number}"], lambda n, t: avances.append((n, t)))
    assert lineas == ["p1", "p2"] and cuenta == "000123456789" and avances == [(1, 2), (2, 2)]