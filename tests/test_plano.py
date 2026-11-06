import pandas as pd
import pytest

from servicios.libro_extracto import escribir_libro
from servicios.plano import GeneradorArchivoPlano, GeneradorExtractosMultiples

NIT = "900111222"
CONSECUTIVOS = {"EG1": 5, "EG2": 7, "RC1": 10, "AJB": 3}
MOVIMIENTOS = [
    ["1/05/2026", "PAGO PROVEEDOR", "900123456", "-1,500,000.00"],
    ["3/05/2026", "IMPTO GOBIERNO 4X1000", "", "-6,000.00"],
    ["10/05/2026", "ABONO INTERESES AHORROS", "", "1,250.50"],
    ["15/05/2026", "PAGO A NOMIN JUAN PEREZ", "", "-900,000.00"],
    ["15/05/2026", "PAGO A NOMIN ANA GOMEZ", "", "-800,000.00"],
    ["20/05/2026", "CONSIGNACION CLIENTE", "800999111", "2,000,000.00"],
    ["2/06/2026", "CUOTA DE MANEJO", "", "-15,000.00"],
    ["9/06/2026", "RETIRO CAJERO", "", "-200,000.00"],
]


@pytest.fixture
def libro(tmp_path):
    df = pd.DataFrame(MOVIMIENTOS, columns=["FECHA", "DESCRIPCIÓN", "DCTO.", "VALOR"])
    return escribir_libro(df, str(tmp_path / "Extracto.xlsx"), cuenta_bancaria="123-456789-01")


@pytest.fixture
def calculado(libro):
    gen = GeneradorArchivoPlano(libro)
    return gen, gen.calcular(gen.preparar(), NIT, CONSECUTIVOS)


def test_clasificacion(libro):
    lin = GeneradorArchivoPlano(libro).preparar()
    movs = lin.drop_duplicates("mov")
    assert movs["clase"].value_counts().to_dict() == {"BANCARIO": 3, "EGRESO": 2, "RECIBO": 1, "NOMINA": 1}
    assert set(lin["lado"]) == {"BANCO", "CONTRA"}


def test_partida_doble_por_comprobante(calculado):
    _, df = calculado
    cuadre = df.groupby(["tipo", "numero"])[["debito", "credito"]].sum()
    assert (cuadre["debito"].round(2) == cuadre["credito"].round(2)).all()


def test_consolidado_bancario_mensual(calculado):
    _, df = calculado
    ajb = df[df["tipo"] == "AJB"]
    assert sorted(ajb["numero"].unique()) == ["0000000003", "0000000004"]
    gmf = ajb[ajb["detalle"].str.startswith("GMF 4X1000")].iloc[0]
    assert (gmf["codpuc"], gmf["dia"], gmf["debito"]) == ("530595001", 31, 6000)
    assert "CUENTA 123-456789-01" in gmf["detalle"]
    intereses = ajb[ajb["detalle"].str.startswith("INTERESES AHORROS")].iloc[0]
    assert (intereses["codpuc"], intereses["credito"]) == ("421005001", 1250.5)
    banco_mayo = ajb[(ajb["codpuc"] == "112005001") & (ajb["mes"] == 5)]
    assert len(banco_mayo) == 1 and banco_mayo["credito"].iloc[0] == pytest.approx(4749.5)


def test_nomina_agrupada_por_dia(calculado):
    _, df = calculado
    nomina = df[df["tipo"] == "EG2"]
    assert nomina["numero"].unique().tolist() == ["0000000007"]
    contra = nomina[nomina["codpuc"] == "250505001"].iloc[0]
    assert contra["debito"] == 1700000 and contra["detalle"] == "PAGO A NOMIN JUAN PEREZ, ANA GOMEZ"


def test_egresos_recibos_cuentas_y_terceros(calculado):
    _, df = calculado
    eg = df[(df["tipo"] == "EG1") & (df["codpuc"] == "220501001")]
    assert eg["tercero"].tolist() == ["900123456", NIT]
    assert eg["numero"].tolist() == ["0000000005", "0000000006"]
    rc = df[(df["tipo"] == "RC1")].set_index("codpuc")
    assert rc.loc["130505001", "tercero"] == "800999111" and rc.loc["130505001", "credito"] == 2000000
    assert rc.loc["112005001", "debito"] == 2000000 and rc.loc["112005001", "tercero"] == NIT


def test_orden_ajb_eg_rc(calculado):
    _, df = calculado
    assert list(dict.fromkeys(df["tipo"].str[:2])) == ["AJ", "EG", "RC"]


def test_edicion_manual_prevalece(libro):
    gen = GeneradorArchivoPlano(libro)
    lin = gen.preparar()
    fila = lin[(lin["clase"] == "RECIBO") & (lin["lado"] == "CONTRA")].index[0]
    lin.loc[fila, ["codpuc_manual", "tercero_manual"]] = ["413595001", "800555444"]
    r = gen.calcular(lin, NIT).loc[fila]
    assert (r["codpuc"], r["tercero"]) == ("413595001", "800555444")


def test_generar_26_columnas(libro, tmp_path, sin_conversion_xls):
    ruta = GeneradorArchivoPlano(libro).generar(nit_empresa=NIT, primer_consecutivo=CONSECUTIVOS,
                                                ruta_destino_final=str(tmp_path / "PLANO.xls"))
    plano = pd.read_excel(ruta, engine="openpyxl", dtype=str)
    assert list(plano.columns) == GeneradorArchivoPlano.COLUMNAS_PLANO
    assert plano["numero"].str.len().eq(10).all()


def test_sin_hoja_datos(tmp_path):
    ruta = tmp_path / "otro.xlsx"
    pd.DataFrame({"A": [1]}).to_excel(ruta, sheet_name="Hoja1", index=False)
    with pytest.raises(ValueError, match="Datos"):
        GeneradorArchivoPlano(str(ruta)).preparar()


def test_columnas_debito_credito(tmp_path):
    ruta = tmp_path / "dc.xlsx"
    pd.DataFrame({"Fecha": ["1/05/2026"], "Concepto": ["PAGO"], "Débitos": ["1.500,00"], "Créditos": [""]}) \
        .to_excel(ruta, sheet_name="Datos", index=False)
    lin = GeneradorArchivoPlano(str(ruta)).preparar()
    assert lin["valor"].iloc[0] == -1500.0 and lin["clase"].iloc[0] == "EGRESO"


def test_multiples_extractos_numeracion_por_archivo(libro, tmp_path):
    otro = escribir_libro(pd.DataFrame([["5/05/2026", "IMPTO GOBIERNO 4X1000", "-400"]],
                                       columns=["FECHA", "DESCRIPCIÓN", "VALOR"]), str(tmp_path / "b.xlsx"))
    gen = GeneradorExtractosMultiples([libro, otro])
    lin = gen.preparar()
    assert lin["archivo"].unique().tolist() == [0, 1]
    assert gen.conteo_comprobantes(lin)["AJB"] == 3