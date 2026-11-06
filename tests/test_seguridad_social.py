import pandas as pd
import pytest

from servicios import aportes_en_linea, seguridad_social
from servicios.seguridad_social import GeneradorPlanillasMultiples, GeneradorSeguridadSocial, generador_para

TEXTO_ARUS = "\n".join([
    "PLANILLA INTEGRADA ARUS", "NIT NI 900111222", "Fecha: 2026-05-10", "Número Planilla: 8812345",
    "Período de Cotización Otros Riesgos: Abril", "Período de Cotización Para Salud: Mayo",
    "N800130907 EPS002 SALUD TOTAL 5 $ 1.234.000 $ 0",
    "N800224808 230301 PORVENIR 5 $ 2.000.000 $ 15.300",
    "N890900841 CCF04 COMFAMA 5 $ 400.000 $ 0",
    "N890903790 14-11 ARL SURA 5 $ 52.200 $ 0",
    "Total a Pagar: $ 3.701.500",
])


@pytest.fixture
def texto_planilla(monkeypatch):
    textos = {}
    monkeypatch.setattr(seguridad_social, "texto_pdf", lambda ruta: textos[ruta])

    def registrar(ruta, texto):
        textos[ruta] = texto
        return ruta

    return registrar


def test_arus_comprobante_ps1(texto_planilla, ruta_administradoras):
    gen = GeneradorSeguridadSocial(texto_planilla("a.pdf", TEXTO_ARUS), ruta_administradoras)
    df = gen.calcular(gen.preparar(), "", {"PS1": 41})
    assert gen.nit_empresa == "900111222" and gen.advertencias == []
    assert df["numero"].unique().tolist() == ["0000000041"]
    assert df["detalle"].iloc[0] == "PAGO SEGURIDAD SOCIAL ABRIL PLANILLA 8812345"
    assert df["codpuc"].tolist() == ["237005001", "238030001", "530520001", "237010001", "237006001", "112005001"]
    assert df["debito"].sum() == df["credito"].sum() == 3701500
    assert df.iloc[-1][["tercero", "nombre"]].tolist() == ["900089104", "ENLACE OPERATIVO"]


def test_arus_total_no_cuadra(texto_planilla, ruta_administradoras):
    texto = TEXTO_ARUS.replace("Total a Pagar: $ 3.701.500", "Total a Pagar: $ 3.800.000")
    with pytest.raises(ValueError, match="no coincide"):
        GeneradorSeguridadSocial(texto_planilla("a.pdf", texto), ruta_administradoras).preparar()


def test_arus_administradora_sin_cuenta(texto_planilla, ruta_administradoras):
    texto = TEXTO_ARUS.replace("N890903790 14-11 ARL SURA", "N860011153 14-23 POSITIVA")
    gen = GeneradorSeguridadSocial(texto_planilla("a.pdf", texto), ruta_administradoras)
    lin = gen.preparar()
    assert "POSITIVA" in gen.advertencias[0] and (lin["codpuc_auto"] == "").sum() == 1


def test_arus_no_es_planilla(texto_planilla, ruta_administradoras):
    with pytest.raises(ValueError, match="Arus"):
        GeneradorSeguridadSocial(texto_planilla("a.pdf", "otro documento"), ruta_administradoras).preparar()


def test_planillas_multiples(texto_planilla, ruta_administradoras):
    otra = TEXTO_ARUS.replace("8812345", "8812399")
    gens = [GeneradorSeguridadSocial(texto_planilla(r, t), ruta_administradoras)
            for r, t in (("a.pdf", TEXTO_ARUS), ("b.pdf", otra))]
    gen = GeneradorPlanillasMultiples(gens)
    df = gen.calcular(gen.preparar(), "", {"PS1": 1})
    assert df["numero"].unique().tolist() == ["0000000001", "0000000002"]


def test_planilla_repetida(texto_planilla, ruta_administradoras):
    gens = [GeneradorSeguridadSocial(texto_planilla(r, TEXTO_ARUS), ruta_administradoras) for r in ("a.pdf", "b.pdf")]
    with pytest.raises(ValueError, match="más de una vez"):
        GeneradorPlanillasMultiples(gens).preparar()


def test_planillas_de_empresas_distintas(texto_planilla, ruta_administradoras):
    otra = TEXTO_ARUS.replace("NIT NI 900111222", "NIT NI 900333444").replace("8812345", "1")
    gens = [GeneradorSeguridadSocial(texto_planilla(r, t), ruta_administradoras)
            for r, t in (("a.pdf", TEXTO_ARUS), ("b.pdf", otra))]
    with pytest.raises(ValueError, match="empresas diferentes"):
        GeneradorPlanillasMultiples(gens).preparar()


def test_generador_para_uno_o_varios():
    class Falso:
        def __init__(self, ruta):
            self.ruta_origen = ruta
    assert isinstance(generador_para(Falso, ["a"]), Falso)
    assert isinstance(generador_para(Falso, ["a", "b"]), GeneradorPlanillasMultiples)


@pytest.fixture
def planilla_aportes(tmp_path):
    filas = [
        ["PLANILLA RESUMEN"], ["EMPRESA SAS NIT: 900111222"],
        ["2026-04 2026-05 1 99887766 E 2026-05-01 2026-05-12 BANCO X 3 $ 3.701.500"],
        ["SALUD TOTAL", "EPS002", "800.130.907", "5", "$ 1.234.000", "$ 0", "$ 0", "$ 1.234.000"],
        ["PORVENIR", "230301", "800224808", "5", "$ 2.000.000", "$ 15.300", "$ 0", "$ 2.015.300"],
        ["COMFAMA", "CCF04", "890900841", "5", "$ 400.000", "$ 0", "$ 0", "$ 400.000"],
        ["ARL SURA", "14-11", "890903790", "5", "$ 52.200", "$ 0", "$ 0", "$ 52.200"],
    ]
    ruta = tmp_path / "SeguridadSocial_ael.xlsx"
    pd.DataFrame(filas).to_excel(ruta, header=False, index=False)
    return str(ruta)


def test_aportes_en_linea_excel(planilla_aportes, ruta_administradoras):
    gen = aportes_en_linea.GeneradorAportesEnLinea(planilla_aportes, ruta_administradoras)
    df = gen.calcular(gen.preparar(), "", {"PS1": 1})
    assert gen.nit_empresa == "900111222"
    assert df["detalle"].iloc[0] == "PAGO SEGURIDAD SOCIAL ABRIL PLANILLA 99887766"
    assert (df["ano"].iloc[0], df["mes"].iloc[0], df["dia"].iloc[0]) == (2026, 5, 12)
    assert df["debito"].sum() == df["credito"].sum() == 3701500
    assert df.iloc[-1]["tercero"] == "900147238"


def test_aportes_en_linea_sin_hoja_resumen(tmp_path, ruta_administradoras):
    ruta = tmp_path / "x.xlsx"
    pd.DataFrame([["otra cosa"]]).to_excel(ruta, header=False, index=False)
    with pytest.raises(ValueError, match="Planilla Resumen"):
        aportes_en_linea.GeneradorAportesEnLinea(str(ruta), ruta_administradoras).preparar()