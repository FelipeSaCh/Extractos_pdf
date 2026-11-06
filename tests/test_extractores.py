import pytest

from extractores import Extraccion, bancolombia, bogota, davivienda, nequi
from nucleo.constantes import DELIMITADOR
from tests.conftest import PaginaFalsa, fila


def _campos(linea):
    return linea.split(DELIMITADOR)


def test_extraccion_dataframe_y_texto():
    e = Extraccion(["A", "B"], ["1|||x", "2|||y", "malformada"])
    assert e.dataframe().values.tolist() == [["1", "x"], ["2", "y"]]
    assert e.texto().splitlines()[0] == "A|||B"


# ---------------------------------------------------------------- Bancolombia
def _pagina_extracto_bancolombia():
    return PaginaFalsa([
        fila(40, (40, "CUENTA DE AHORROS No. 123-456789-01")),
        fila(80, (40, "FECHA"), (130, "DESCRIPCIÓN"), (300, "SUCURSAL"), (420, "DCTO."), (480, "VALOR"), (550, "SALDO")),
        fila(100, (40, "1/05"), (130, "PAGO PROVEEDOR"), (300, "MEDELLIN"), (420, "900123456"),
             (480, "-1,500,000.00"), (550, "8,500,000.00")),
        fila(115, (40, "3/05"), (130, "IMPTO GOBIERNO 4X1000"), (480, "-6,000.00"), (550, "8,494,000.00")),
        fila(760, (270, "Página 1 de 1")),
    ])


def test_bancolombia_extracto(pdf_falso):
    ruta = pdf_falso("Extracto_x.pdf", [_pagina_extracto_bancolombia()])
    e = bancolombia.extraer(ruta, tipo_documento="Extracto")
    assert e.columnas == bancolombia.COLUMNAS_EXTRACTO
    assert e.cuenta == "123-456789-01"
    assert [_campos(l) for l in e.lineas] == [
        ["1/05", "PAGO PROVEEDOR", "MEDELLIN", "900123456", "-1,500,000.00", "8,500,000.00"],
        ["3/05", "IMPTO GOBIERNO 4X1000", "", "", "-6,000.00", "8,494,000.00"],
    ]


def test_bancolombia_movimientos_sociedad(pdf_falso):
    pagina = PaginaFalsa([
        fila(80, (40, "FECHA"), (150, "DESCRIPCIÓN"), (320, "REFERENCIA 1")),
        fila(100, (40, "2026/05/02"), (150, "PAGO PROV"), (240, "MED"), (320, "900123456"), (400, "X1"),
             (470, "77"), (540, "-50,000.00")),
        fila(112, (150, "CONTINUA")),
    ])
    e = bancolombia.extraer(pdf_falso("cualquier.pdf", [pagina]), tipo_documento="MovimientoSOC")
    assert [_campos(l) for l in e.lineas] == [
        ["2026/05/02", "PAGO PROV CONTINUA", "MED", "900123456", "X1", "77", "-50,000.00"]]


def test_bancolombia_persona_natural_montos_con_punto_de_miles(pdf_falso):
    pagina = PaginaFalsa([
        fila(80, (40, "Fecha"), (200, "Descripción")),
        fila(100, (40, "5 ene 2026"), (200, "Transferencia"), (400, "12345678"), (520, "-$50.000")),
        fila(115, (40, "7 ene 2026"), (200, "Abono"), (520, "$1.234.567,89")),
    ])
    e = bancolombia.extraer(pdf_falso("x.pdf", [pagina]), tipo_documento="MovimientosPNatural")
    assert [_campos(l)[-1] for l in e.lineas] == ["-50000.00", "1234567.89"]
    assert _campos(e.lineas[0])[0] == "5 ene 2026"


@pytest.mark.parametrize("nombre, tipo", [
    ("Extracto_abril.pdf", "Extracto"),
    ("MovimientosPNatural_x.pdf", "MovimientosPNatural"),
    ("movimiento_soc.pdf", "MovimientoSOC"),
])
def test_bancolombia_tipo_por_nombre(nombre, tipo):
    assert bancolombia.tipo_por_nombre(nombre) == tipo


def test_bancolombia_nombre_invalido_sin_tipo():
    with pytest.raises(ValueError):
        bancolombia.tipo_por_nombre("abril.pdf")


def test_bancolombia_boton_prevalece_sobre_nombre(pdf_falso):
    ruta = pdf_falso("movimientos_abril.pdf", [_pagina_extracto_bancolombia()])
    assert bancolombia.extraer(ruta, tipo_documento="Extracto").columnas == bancolombia.COLUMNAS_EXTRACTO


# ---------------------------------------------------------------- Bogotá / Nequi
def test_bogota(pdf_falso):
    pagina = PaginaFalsa([
        fila(80, (40, "FECHA"), (150, "DESCRIPCIÓN")),
        fila(100, (40, "02/05"), (75, "123"), (150, "PAGO PROVEEDOR"), (290, "BOGOTA"), (350, "OF1"),
             (420, "55"), (470, "-100,000.00"), (555, "900,000.00")),
        fila(112, (150, "DETALLE")),
        fila(130, (40, "FIN MOVIMIENTOS")),
        fila(150, (40, "09/05"), (150, "NO DEBE APARECER"), (470, "1.00")),
    ])
    e = bogota.extraer(pdf_falso("b.pdf", [pagina]))
    assert [_campos(l) for l in e.lineas] == [
        ["02/05", "123", "PAGO PROVEEDOR DETALLE", "BOGOTA", "OF1", "55", "-100,000.00", "900,000.00"]]


def test_nequi(pdf_falso):
    pagina = PaginaFalsa([
        fila(80, (40, "Fecha"), (200, "Descripción"), (420, "Valor"), (520, "Saldo")),
        fila(100, (40, "01/05/2026"), (200, "Recibí de Ana"), (420, "$50,000.00"), (520, "$150,000.00")),
        fila(115, (40, "03/05/2026"), (200, "Pago tienda"), (420, "-$20,000.00"), (520, "$130,000.00")),
        fila(140, (40, "Vigilado Superintendencia")),
    ])
    e = nequi.extraer(pdf_falso("n.pdf", [pagina]))
    assert [_campos(l) for l in e.lineas] == [
        ["01/05/2026", "Recibí de Ana", "$50,000.00", "$150,000.00"],
        ["03/05/2026", "Pago tienda", "-$20,000.00", "$130,000.00"],
    ]


# ---------------------------------------------------------------- Davivienda
def _pagina_davivienda_comun():
    return PaginaFalsa([
        fila(40, (40, "SALDO ANTERIOR $1,000,000.00"), (300, "NUEVO SALDO $1,140,000.00")),
        fila(55, (40, "MAYO / 2026")),
        fila(80, (30, "FECHA"), (80, "VALOR"), (180, "DOC"), (260, "CLASE DE MOVIMIENTO"), (460, "OFICINA")),
        fila(100, (30, "02"), (45, "05"), (80, "$200,000.00+"), (190, "111"), (260, "CONSIGNACION"), (460, "CENTRO")),
        fila(115, (30, "04"), (45, "05"), (80, "$60,000.00-"), (190, "222"), (260, "PAGO SERVICIO"), (460, "WEB")),
        fila(127, (260, "ADICIONAL")),
        fila(150, (40, "LA IMPORTANCIA DEL AHORRO")),
    ])


def test_davivienda_comun(pdf_falso):
    e = davivienda.extraer(pdf_falso("d.pdf", [_pagina_davivienda_comun()]))
    assert (e.saldo_anterior, e.saldo_actual) == (1000000.0, 1140000.0)
    assert [_campos(l) for l in e.lineas] == [
        ["02/05/2026", "", "CONSIGNACION", "", "CENTRO", "111", "200000.00", ""],
        ["04/05/2026", "", "PAGO SERVICIO ADICIONAL", "", "WEB", "222", "-60000.00", ""],
    ]


def test_davivienda_damas(pdf_falso):
    pagina = PaginaFalsa([
        fila(40, (40, "SALDO ANTERIOR $500,000.00"), (300, "NUEVO SALDO $530,000.00")),
        fila(55, (40, "JUNIO / 2026")),
        fila(80, (30, "FECHA"), (150, "DESCRIPCIÓN"), (420, "DÉBITO"), (500, "CRÉDITO")),
        fila(100, (30, "03"), (45, "06"), (70, "0021"), (150, "ABONO NOMINA"), (380, "998877"),
             (430, "$0.00"), (510, "$50,000.00")),
        fila(112, (150, "EMPRESA X")),
        fila(130, (30, "08"), (45, "06"), (150, "COMPRA SUPERMERCADO"), (430, "$20,000.00"), (510, "$0.00")),
    ])
    e = davivienda.extraer(pdf_falso("d.pdf", [pagina]))
    assert [(_campos(l)[0], _campos(l)[2], _campos(l)[-2]) for l in e.lineas] == [
        ("03/06/2026", "ABONO NOMINA EMPRESA X", "50000.00"),
        ("08/06/2026", "COMPRA SUPERMERCADO", "-20000.00"),
    ]


def test_davivienda_tipo_equivocado(pdf_falso):
    ruta = pdf_falso("d.pdf", [_pagina_davivienda_comun()])
    assert davivienda.detectar_tipo_cuenta(ruta) == davivienda.TIPO_COMUN
    with pytest.raises(ValueError, match="DAMAS"):
        davivienda.extraer(ruta, tipo_cuenta=davivienda.TIPO_DAMAS)