import datetime
import re

import numpy as np
import pandas as pd
import pdfplumber

from script_seguridad_social import (
    GeneradorSeguridadSocial,
    RUTA_ADMINISTRADORAS,
    _cargar_administradoras,
    _normalizar,
)

PATRON_LIQ = re.compile(
    r"(?P<ap>\d{4})[-/](?P<mp>\d{1,2})(?:[-/]\d{1,2})?\s+\d{4}[-/]\d{1,2}(?:[-/]\d{1,2})?\s+\d+\s+(?P<planilla>\d+)\s+\S+\s+"
    r"\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+(?P<ano>\d{4})[-/](?P<mes>\d{1,2})[-/](?P<dia>\d{1,2})\s+.*?\s+(?:\d+\s+)?\$?\s*(?P<valor>[\d.,]+)\s*$",
    re.M,
)
PATRON_NIT = re.compile(r"\bNIT\s*:?\s*(\d{5,15})")
_E = r"[ \t]+"
PATRON_MORA = re.compile(
    rf"^[ \t]*(?P<adm>\S.*?){_E}(?P<cod>\S+){_E}(?P<nit>\d{{1,3}}(?:[.,]\d{{3}}){{2,}}|\d{{6,15}})(?:{_E}\d)?{_E}\d+"
    rf"{_E}\$?[ \t]*[\d.,]+{_E}\$?[ \t]*(?P<mora>[\d.,]+){_E}\$?[ \t]*-?[\d.,]+{_E}\$?[ \t]*[\d.,]+[ \t]*$",
    re.M,
)

ADMINISTRADORAS_PILA = {
    "25-14": ("900336004", "COLPENSIONES"),
    "230201": ("800138188", "PROTECCION"),
    "230301": ("800144331", "PORVENIR"),
    "230901": ("800148514", "SKANDIA"),
    "231001": ("800149496", "COLFONDOS"),
    "EPS002": ("800130907", "SALUD TOTAL"),
    "EPS005": ("800251440", "SANITAS"),
    "EPS008": ("860066942", "COMPENSAR"),
    "EPS010": ("800088702", "EPS SURA"),
    "EPS037": ("900156264", "NUEVA EPS"),
    "CCF04": ("890900841", "COMFAMA"),
    "CCF05": ("890900842", "COMFENALCO ANTIOQUIA"),
    "14-23": ("860011153", "POSITIVA"),
    "14-28": ("800256161", "ARL SURA"),
    "SENA": ("899999034", "SENA"),
    "ICBF": ("899999239", "ICBF"),
}


def _celda(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (datetime.date, pd.Timestamp)):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def _clave(texto):
    return _normalizar(texto).replace(" ", "")


def _codigo(texto):
    return re.sub(r"\s", "", str(texto or "")).upper()


def _valor(texto):
    txt = re.sub(r"[^\d.,]", "", str(texto or ""))
    txt = re.sub(r"[.,](?=\d{3}(?!\d))", "", txt).replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return 0.0


class GeneradorAportesEnLinea(GeneradorSeguridadSocial):
    NIT_OPERADOR = "900147238"
    NOMBRE_OPERADOR = "APORTES EN LINEA"
    GRUPOS = ("PENSION", "SALUD", "CCF", "RIESGOS", "PARAFISCALES")
    EXTENSIONES_EXCEL = (".xls", ".xlsx", ".xlsm")
    MARCAS_HOJA = ("PLANILLA RESUMEN", "LIQUIDACION DETALLADA")

    def __init__(self, ruta_origen, ruta_administradoras=RUTA_ADMINISTRADORAS):
        super().__init__(ruta_origen, ruta_administradoras)

    def _leer_excel(self):
        hojas = pd.read_excel(self.ruta_origen, sheet_name=None, header=None, dtype=object)
        libros = [[[_celda(v) for v in f] for f in df.itertuples(index=False)] for df in hojas.values()]
        textos = ["\n".join(" ".join(c.replace("\n", " ") for c in f if c) for f in l) for l in libros]
        for marca in self.MARCAS_HOJA:
            idx = next((i for i, l in enumerate(libros) if any(marca in _normalizar(c) for f in l for c in f)), None)
            if idx is not None:
                return "\n".join([textos[idx]] + [t for i, t in enumerate(textos) if i != idx]), [libros[idx]]
        raise ValueError("No se encontró la hoja 'Planilla Resumen' en el Excel.")

    @staticmethod
    def _moras(texto):
        moras = {}
        for m in PATRON_MORA.finditer(texto):
            monto = _valor(m["mora"])
            if not monto:
                continue
            nit = re.sub(r"\D", "", m["nit"])
            nombre, previo = moras.get(nit, (m["adm"].strip(), 0.0))
            moras[nit] = (nombre, previo + monto)
        return moras

    def _leer_pdf_tablas(self):
        textos, bloques = [], []
        with pdfplumber.open(self.ruta_origen) as pdf:
            for pagina in pdf.pages:
                p = pagina.filter(lambda o: o.get("object_type") != "char" or o.get("upright", True))
                textos.append(p.extract_text() or "")
                bloques += [[[str(c or "").strip() for c in f] for f in t] for t in p.extract_tables()]
        return "\n".join(textos), bloques

    def _leer_origen(self):
        if self.ruta_origen.lower().endswith(self.EXTENSIONES_EXCEL):
            return self._leer_excel()
        return self._leer_pdf_tablas()

    def _layout(self, filas):
        for i, fila in enumerate(filas):
            claves = [_clave(c) for c in fila]
            if not all(g in claves for g in self.GRUPOS):
                continue
            j = next((k for k in range(i + 1, min(i + 4, len(filas))) if "IBC" in map(_clave, filas[k])), None)
            if j is None:
                continue
            inicios = sorted((claves.index(g), g) for g in self.GRUPOS)
            sub = [(c, _clave(t)) for c, t in enumerate(filas[j]) if t]
            limites = [c for c, _ in sub[1:]] + [len(filas[j])]
            campos = {}
            for (c, t), fin in zip(sub, limites):
                grupo = next((g for s, g in reversed(inicios) if c >= s), None)
                if grupo is None:
                    if t == "NO":
                        campos.setdefault("NO", (c, fin))
                elif t.startswith("COD"):
                    campos.setdefault((grupo, "COD"), (c, fin))
                elif t == "APORTE":
                    campos.setdefault((grupo, "APORTE"), (c, fin))
            faltan = [g for g in self.GRUPOS if (g, "APORTE") not in campos] + (["No."] if "NO" not in campos else [])
            if faltan:
                raise ValueError(f"No se identificaron las columnas de: {', '.join(faltan)}.")
            return j + 1, campos
        return None, None

    @staticmethod
    def _en(fila, span):
        return next((c for c in fila[span[0]:span[1]] if c), "")

    def _aportes(self, bloques):
        aportes, campos = {}, None
        for filas in bloques:
            inicio, nuevos = self._layout(filas)
            campos = nuevos or campos
            if not campos:
                continue
            for fila in filas[inicio or 0:]:
                if not self._en(fila, campos["NO"]).strip().isdigit():
                    continue
                for g in self.GRUPOS:
                    valor = _valor(self._en(fila, campos[(g, "APORTE")]))
                    if valor:
                        cod = _codigo(self._en(fila, campos[(g, "COD")])) if (g, "COD") in campos else ""
                        aportes[(g, cod)] = aportes.get((g, cod), 0.0) + valor
        return aportes

    def _administradora(self, codigo):
        fila = self._adm[(self._adm["codigo"] == codigo) & (self._adm["nit"] != "")]
        if not fila.empty:
            return fila["nit"].iloc[0], fila["nombre"].iloc[0]
        return ADMINISTRADORAS_PILA.get(codigo, ("", ""))

    def preparar(self):
        texto, bloques = self._leer_origen()
        liq = PATRON_LIQ.search(texto)
        if not liq:
            raise ValueError("No se encontraron los datos generales de la liquidación. Verifique que sea una Planilla Resumen de Aportes en Línea.")

        nit_emp = PATRON_NIT.search(texto)
        self.nit_empresa = nit_emp.group(1) if nit_emp else ""
        self._adm = _cargar_administradoras(self.ruta_administradoras)
        aportes = self._aportes(bloques)
        if not aportes:
            raise ValueError("No se encontraron aportes en la liquidación detallada.")

        filas, sin_cuenta = [], []
        for (grupo, cod), valor in sorted(aportes.items()):
            if grupo == "PARAFISCALES":
                sena = round(valor * 2 / 5, -2)
                partes = [("SENA", sena), ("ICBF", valor - sena)]
            else:
                partes = [(cod, valor)]
            for codigo, monto in partes:
                nit, nombre = self._administradora(codigo)
                cuenta = self._buscar_cuenta(codigo, nit, nombre)
                if not cuenta:
                    sin_cuenta.append(f"{grupo} {codigo} {nombre}".strip())
                filas.append(self._fila(cuenta, nit, nombre or codigo, debito=monto))

        moras = self._moras(texto)
        for nit, (nombre, monto) in moras.items():
            filas.append(self._fila(self.CODPUC_INTERESES_MORA, nit, nombre, debito=monto))

        total = sum(aportes.values()) + sum(m for _, m in moras.values())
        pagado = _valor(liq["valor"])
        mora = round(pagado - total, 2)
        if mora < -1:
            raise ValueError(f"La suma de aportes e intereses ({total:,.0f}) supera el valor pagado ({pagado:,.0f}).")
        if mora > 1:
            filas.append(self._fila(self.CODPUC_INTERESES_MORA, self.NIT_OPERADOR, self.NOMBRE_OPERADOR, debito=mora))
            filas.append(self._fila(self.CODPUC_INTERESES_MORA, self.NIT_OPERADOR, self.NOMBRE_OPERADOR, debito=mora))

        self.advertencias = (
            ["Sin tercero o cuenta contable en administradoras.json:\n- " + "\n- ".join(sin_cuenta)] if sin_cuenta else []
        )

        detalle = f"PAGO SEGURIDAD SOCIAL {self.MESES[int(liq['mp']) - 1]} PLANILLA {liq['planilla']}"
        return pd.DataFrame(self._con_contrapartida_total(filas)).assign(
            mov=0, clase=self.CLASE_SS, tipo=self.COMPROBANTES_PREDETERMINADOS[self.CLASE_SS],
            ano=int(liq["ano"]), mes=int(liq["mes"]), dia=int(liq["dia"]), detalle=detalle,
            codpuc_manual="", tercero_manual="", docrela="",
        )