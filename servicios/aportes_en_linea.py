"""Planilla Resumen de Aportes en Línea (PDF o Excel) → comprobante PS1."""
import datetime
import re

import numpy as np
import pandas as pd

from infraestructura.lectura_pdf import texto_y_tablas
from nucleo.config import RUTA_ADMINISTRADORAS
from nucleo.numeros import a_float
from nucleo.texto import clave as _clave, codigo as _codigo, normalizar
from servicios import seguridad_social
from servicios.seguridad_social import ARL, PENSION, SALUD, TIPO_EMPLEADOS, TOLERANCIA, GeneradorSeguridadSocial

PATRON_LIQ = re.compile(
    r"(?P<ap>\d{4})[-/](?P<mp>\d{1,2})(?:[-/]\d{1,2})?\s+\d{4}[-/]\d{1,2}(?:[-/]\d{1,2})?\s+\d+\s+(?P<planilla>\d+)\s+\S+\s+"
    r"\d{4}[-/]\d{1,2}[-/]\d{1,2}\s+(?P<ano>\d{4})[-/](?P<mes>\d{1,2})[-/](?P<dia>\d{1,2})\s+.*?\s+(?:\d+\s+)?\$?\s*(?P<valor>[\d.,]+)\s*$",
    re.M,
)
PATRON_NIT = re.compile(r"\bNIT\s*:?\s*(\d{5,15})")
_E = r"[ \t]+"
_M = r"\$?[ \t]*(-?[\d.,]+)"
PATRON_RESUMEN = re.compile(
    rf"^[ \t]*(?P<adm>\S.*?){_E}(?P<cod>\S+){_E}(?P<nit>\d{{1,3}}(?:[.,]\d{{3}}){{2,}}|\d{{6,15}})(?:{_E}\d)?{_E}\d+"
    rf"{_E}{_M.replace('(', '(?P<val>', 1)}{_E}{_M.replace('(', '(?P<mora>', 1)}"
    rf"{_E}{_M.replace('(', '(?P<saldo>', 1)}{_E}{_M.replace('(', '(?P<pagar>', 1)}[ \t]*$",
    re.M,
)

ADMINISTRADORAS_PILA = {
    "25-14": ("900336004", "COLPENSIONES"),
    "230201": ("800229739", "PROTECCION"),
    "230301": ("800224808", "PORVENIR"),
    "230901": ("800148514", "SKANDIA"),
    "231001": ("800149496", "COLFONDOS"),
    "EPS002": ("800130907", "SALUD TOTAL"),
    "EPS005": ("800251440", "SANITAS"),
    "EPS008": ("860066942", "COMPENSAR"),
    "EPS010": ("800088702", "EPS SURA"),
    "EPS037": ("900156264", "NUEVA EPS"),
    "EPS040": ("900604350", "SAVIA SALUD"),
    "ESSC07": ("806008394", "MUTUAL SER"),
    "CCF04": ("890900841", "COMFAMA"),
    "CCF05": ("890102044", "CAJACOPI ATLANTICO"),
    "CCF56": ("890303093", "COMFENALCO VALLE"),
    "14-23": ("860011153", "POSITIVA"),
    "14-11": ("890903790", "ARL SURA"),
    "14-28": ("800256161", "ARL SURA"),
    "SENA": ("899999034", "SENA"),
    "ICBF": ("899999239", "ICBF"),
}
PROPORCION_SENA = 2 / 5
OPCIONES = seguridad_social.OPCIONES


def _celda(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (datetime.date, pd.Timestamp)):
        return v.strftime("%Y/%m/%d")
    if isinstance(v, (float, np.floating)) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


class GeneradorAportesEnLinea(GeneradorSeguridadSocial):
    NIT_OPERADOR = "900147238"
    NOMBRE_OPERADOR = "APORTES EN LINEA"
    CODPUC_SALDOS = ""
    GRUPOS = ("PENSION", "SALUD", "CCF", "RIESGOS", "PARAFISCALES")
    SUBSISTEMA_GRUPO = {"PENSION": PENSION, "SALUD": SALUD, "RIESGOS": ARL}
    EXTENSIONES_EXCEL = (".xls", ".xlsx", ".xlsm")
    MARCAS_HOJA = ("PLANILLA RESUMEN", "LIQUIDACION DETALLADA")
    MSG_SIN_CUENTA = "Sin tercero o cuenta contable en administradoras.json"

    def __init__(self, ruta_origen, ruta_administradoras=RUTA_ADMINISTRADORAS, tipo_planilla=TIPO_EMPLEADOS):
        super().__init__(ruta_origen, ruta_administradoras, tipo_planilla)

    def _leer_excel(self):
        hojas = pd.read_excel(self.ruta_origen, sheet_name=None, header=None, dtype=object)
        libros = [[[_celda(v) for v in f] for f in df.itertuples(index=False)] for df in hojas.values()]
        textos = ["\n".join(" ".join(c.replace("\n", " ") for c in f if c) for f in l) for l in libros]
        for marca in self.MARCAS_HOJA:
            idx = next((i for i, l in enumerate(libros) if any(marca in normalizar(c) for f in l for c in f)), None)
            if idx is not None:
                return "\n".join([textos[idx]] + [t for i, t in enumerate(textos) if i != idx]), [libros[idx]]
        raise ValueError("No se encontró la hoja 'Planilla Resumen' en el Excel.")

    def _leer_origen(self):
        if self.ruta_origen.lower().endswith(self.EXTENSIONES_EXCEL):
            return self._leer_excel()
        return texto_y_tablas(self.ruta_origen)

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
                    valor = abs(a_float(self._en(fila, campos[(g, "APORTE")])))
                    if valor:
                        cod = _codigo(self._en(fila, campos[(g, "COD")])) if (g, "COD") in campos else ""
                        aportes[(g, cod)] = aportes.get((g, cod), 0.0) + valor
        return aportes

    def _administradora(self, codigo):
        fila = self._adm[(self._adm["codigo"] == codigo) & (self._adm["nit"] != "")]
        if not fila.empty:
            return fila["nit"].iloc[0], fila["nombre"].iloc[0]
        return ADMINISTRADORAS_PILA.get(codigo, ("", ""))

    @staticmethod
    def _resumen(texto):
        filas = {}
        for m in PATRON_RESUMEN.finditer(texto):
            nit = re.sub(r"\D", "", m["nit"])
            clave = (_codigo(m["cod"]), nit)
            previo = filas.get(clave, {"nombre": m["adm"].strip(), "val": 0.0, "mora": 0.0, "saldo": 0.0, "pagar": 0.0})
            for k in ("val", "mora", "saldo", "pagar"):
                previo[k] += abs(a_float(m[k]))
            filas[clave] = previo
        return filas

    def _filas_detalle(self, bloques, sin_cuenta):
        aportes = self._aportes(bloques)
        if not aportes:
            raise ValueError("No se encontraron aportes en la planilla.")
        filas = []
        for (grupo, cod), valor in sorted(aportes.items()):
            if grupo == "PARAFISCALES":
                sena = round(valor * PROPORCION_SENA, -2)
                partes = [("SENA", sena), ("ICBF", valor - sena)]
            else:
                partes = [(cod, valor)]
            for cod_adm, monto in partes:
                nit, nombre = self._administradora(cod_adm)
                cuenta = self._cuenta(cod_adm, nit, nombre, self.SUBSISTEMA_GRUPO.get(grupo))
                if not cuenta:
                    sin_cuenta.append(f"{grupo} {cod_adm} {nombre}".strip())
                filas.append(self._fila(cuenta, nit, nombre or cod_adm, debito=monto))
        return filas

    def _filas_resumen(self, resumen, sin_cuenta):
        filas = []
        for (cod, nit), r in resumen.items():
            if r["val"]:
                cuenta = self._cuenta(cod, nit, r["nombre"])
                if not cuenta:
                    sin_cuenta.append(f"{cod} {r['nombre']}")
                filas.append(self._fila(cuenta, nit, r["nombre"], debito=r["val"]))
            if r["mora"]:
                filas.append(self._fila(self.CODPUC_INTERESES_MORA, nit, r["nombre"], debito=r["mora"]))
            if r["saldo"]:
                filas.append(self._fila(self.CODPUC_SALDOS, nit, r["nombre"], credito=r["saldo"]))
        return filas

    def preparar(self):
        texto, bloques = self._leer_origen()
        liq = PATRON_LIQ.search(texto)
        if not liq:
            raise ValueError("No se encontraron los datos generales de la liquidación. "
                             "Verifique que sea una Planilla Resumen de Aportes en Línea.")

        nit_emp = PATRON_NIT.search(texto)
        self.nit_empresa = nit_emp.group(1) if nit_emp else ""
        self._cargar_administradoras()

        sin_cuenta = []
        resumen = self._resumen(texto)
        filas = self._filas_resumen(resumen, sin_cuenta) if resumen else self._filas_detalle(bloques, sin_cuenta)

        total = round(self._suma(filas, "debito") - self._suma(filas, "credito"), 2)
        pagado = a_float(liq["valor"])
        diferencia = round(pagado - total, 2)
        if resumen and abs(diferencia) > TOLERANCIA:
            raise ValueError(f"El resumen de pago ({total:,.0f}) no coincide con el valor pagado ({pagado:,.0f}).")
        if diferencia < -TOLERANCIA:
            raise ValueError(f"La suma de aportes ({total:,.0f}) supera el valor pagado ({pagado:,.0f}).")
        if diferencia > TOLERANCIA:
            filas.append(self._fila(self.CODPUC_INTERESES_MORA, self.NIT_OPERADOR, self.NOMBRE_OPERADOR, debito=diferencia))

        self._advertir_sin_cuenta(sin_cuenta)
        if any(pd.notna(f["credito"]) and not f["codpuc_auto"] for f in filas):
            self.advertencias.append("La planilla descuenta saldos/incapacidades: "
                                     "asigne la cuenta por cobrar en las líneas crédito sin cuenta.")

        detalle = f"PAGO SEGURIDAD SOCIAL {self.MESES[int(liq['mp']) - 1]} PLANILLA {liq['planilla']}"
        filas.append(self._contrapartida(pagado))
        return self._comprobante(filas, int(liq["ano"]), int(liq["mes"]), int(liq["dia"]), detalle)


Generador = GeneradorAportesEnLinea