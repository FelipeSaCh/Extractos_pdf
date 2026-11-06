"""Planillas PILA de ARUS (Enlace Operativo) → comprobante PS1."""
import os
import re

import numpy as np
import pandas as pd

from infraestructura import administradoras
from infraestructura.lectura_pdf import texto_pdf
from nucleo.config import RUTA_ADMINISTRADORAS
from nucleo.numeros import a_float
from nucleo.texto import codigo, normalizar
from servicios.plano import GeneradorArchivoPlano

PATRON_FILA = re.compile(
    r"^(?:N(?P<nit>\d{6,15})\s+)?(?P<cod>\S+)(?:\s+(?P<adm>.*?))?\s+(?P<afi>\d+)"
    r"\s+\$\s*(?P<val>[\d.,]+)\s+\$\s*(?P<mora>[\d.,]+)\s*$"
)
PATRON_FECHA = re.compile(r"Fecha\s*:?\s*(\d{4})-(\d{1,2})-(\d{1,2})")
PATRON_PLANILLA = re.compile(r"N[úu]mero\s+Planilla\s*:?\s*(\d+)", re.I)
PATRON_NIT_EMPRESA = re.compile(r"^NIT\s+[A-Z]{1,2}\s*(\d{5,15})", re.M)
PATRON_TOTAL = re.compile(r"Total\s+a\s+Pagar\s*:?\s*\$\s*([\d.,]+)", re.I)
PATRON_PERIODO = re.compile(
    r"Per[ií]odo\s+de\s+Cotizaci[oó]n\s+(?P<tipo>Otros\s+Riesgos|Para\s+Salud)\s*:?\s*(?P<mes>[A-Za-zÁÉÍÓÚáéíóúÑñ]+)",
    re.I,
)
TOLERANCIA = 1

TIPO_EMPLEADOS = "EMPLEADOS"
TIPO_INDEPENDIENTE = "INDEPENDIENTE"
TIPOS_PLANILLA = {
    TIPO_EMPLEADOS: "Planilla de empleados",
    TIPO_INDEPENDIENTE: "Planilla de independiente",
}
OPCIONES = {"tipo_planilla": ("Tipo de planilla", "Seleccione el tipo de planilla de seguridad social:",
                              TIPOS_PLANILLA, None)}

SALUD, PENSION, ARL, CCF, PARAFISCALES = "SALUD", "PENSION", "ARL", "CCF", "PARAFISCALES"
_SUBSISTEMA_POR_CODIGO = (
    (re.compile(r"^(EPS|ESS|EAS|EPSI|EPSIC|RES)"), SALUD),
    (re.compile(r"^(23\d{4}|25-14)$"), PENSION),
    (re.compile(r"^14-\d+$"), ARL),
    (re.compile(r"^CCF"), CCF),
    (re.compile(r"^(SENA|ICBF)$"), PARAFISCALES),
)
_SUBSISTEMA_POR_CATEGORIA = (("EPS", SALUD), ("SALUD", SALUD), ("PENSION", PENSION), ("ARL", ARL),
                             ("RIESGO", ARL), ("CCF", CCF), ("CAJA", CCF), ("PARAFISCAL", PARAFISCALES))


class GeneradorSeguridadSocial(GeneradorArchivoPlano):
    CLASE_SS = "SEGURIDAD_SOCIAL"
    COMPROBANTES_PREDETERMINADOS = {CLASE_SS: "PS1"}
    CLAVE_CONSECUTIVO = "consecutivo_seguridad_social"
    NIT_ARUS = "900089104"
    NIT_OPERADOR = NIT_ARUS
    NOMBRE_OPERADOR = "ENLACE OPERATIVO"
    CODPUC_INTERESES_MORA = "530520001"
    CODPUC_CONTRAPARTIDA = "112005001"
    MSG_SIN_CUENTA = "Sin cuenta contable en administradoras.json"
    CUENTAS_INDEPENDIENTE = {SALUD: "519605001", PENSION: "519605002", ARL: "519605003"}

    def __init__(self, ruta_pdf, ruta_administradoras=RUTA_ADMINISTRADORAS, tipo_planilla=TIPO_EMPLEADOS):
        super().__init__(ruta_pdf)
        self.ruta_administradoras = ruta_administradoras
        self.tipo_planilla = tipo_planilla or TIPO_EMPLEADOS
        self.nit_empresa = ""
        self.advertencias = []
        self._adm = None

    def _cargar_administradoras(self):
        self._adm = administradoras.tabla(self.ruta_administradoras)

    def _buscar_cuenta(self, cod, nit, nombre):
        adm = self._adm[self._adm["cuenta"] != ""]
        cod = codigo(cod)
        if cod:
            cuentas = adm.loc[adm["codigo"] == cod, "cuenta"].unique()
            if len(cuentas) == 1:
                return cuentas[0]
        cand = adm[adm["nit"] == nit] if nit else adm.iloc[0:0]
        if cand["cuenta"].nunique() > 1:
            cand = cand[~cand["categoria"].str.contains("CESANT", na=False)]
        if cand["cuenta"].nunique() == 1:
            return cand["cuenta"].iloc[0]
        nom = normalizar(nombre)
        if not nom:
            return ""
        base = cand if not cand.empty else adm
        for filtro in (
            base["nombre"] == nom,
            base["nombre"].map(lambda n: bool(n) and (n.startswith(nom) or nom.startswith(n))),
        ):
            cuentas = base.loc[filtro, "cuenta"].unique()
            if len(cuentas) == 1:
                return cuentas[0]
        return ""

    def _subsistema(self, cod, nit):
        cod = codigo(cod)
        for patron, subsistema in _SUBSISTEMA_POR_CODIGO:
            if patron.match(cod):
                return subsistema
        categorias = self._adm.loc[(self._adm["nit"] == nit) & (nit != ""), "categoria"]
        return next((s for cat in categorias for clave, s in _SUBSISTEMA_POR_CATEGORIA if clave in cat), "")

    def _cuenta(self, cod, nit, nombre, subsistema=None):
        """Cuenta del aporte: en planillas de independiente, salud/pensión/ARL van a las cuentas 5196."""
        if self.tipo_planilla == TIPO_INDEPENDIENTE:
            cuenta = self.CUENTAS_INDEPENDIENTE.get(subsistema or self._subsistema(cod, nit))
            if cuenta:
                return cuenta
        return self._buscar_cuenta(cod, nit, nombre)

    def _mes_periodo(self, texto, mes_fecha):
        periodos = {normalizar(m["tipo"]): normalizar(m["mes"]) for m in PATRON_PERIODO.finditer(texto)}
        nombre = periodos.get("OTROS RIESGOS") or periodos.get("PARA SALUD") or ""
        return next((m for m in self.MESES if nombre and m[:3] == nombre[:3]), self.MESES[mes_fecha - 1])

    @staticmethod
    def _fila(cuenta, tercero, nombre, debito=np.nan, credito=np.nan):
        return {"codpuc_auto": cuenta, "tercero_ref": tercero, "nombre": nombre, "debito": debito, "credito": credito}

    @staticmethod
    def _suma(filas, campo):
        return sum(f[campo] for f in filas if pd.notna(f[campo]))

    def _contrapartida(self, valor):
        return self._fila(self.CODPUC_CONTRAPARTIDA, self.NIT_OPERADOR, self.NOMBRE_OPERADOR, credito=valor)

    def _advertir_sin_cuenta(self, sin_cuenta):
        self.advertencias = [f"{self.MSG_SIN_CUENTA}:\n- " + "\n- ".join(sin_cuenta)] if sin_cuenta else []

    def _comprobante(self, filas, ano, mes, dia, detalle):
        return pd.DataFrame(filas).assign(
            mov=0, clase=self.CLASE_SS, tipo=self.COMPROBANTES_PREDETERMINADOS[self.CLASE_SS],
            ano=ano, mes=mes, dia=dia, detalle=detalle,
            codpuc_manual="", tercero_manual="", docrela="",
        )

    def preparar(self):
        texto = texto_pdf(self.ruta_origen)
        fecha = PATRON_FECHA.search(texto)
        planilla = PATRON_PLANILLA.search(texto)
        if not fecha or not planilla:
            raise ValueError("No se encontró la Fecha o el Número de Planilla. "
                             "Verifique que sea una planilla de Arus (Enlace Operativo).")

        ano, mes, dia = map(int, fecha.groups())
        nit_emp = PATRON_NIT_EMPRESA.search(texto)
        self.nit_empresa = nit_emp.group(1) if nit_emp else ""
        detalle = f"PAGO SEGURIDAD SOCIAL {self._mes_periodo(texto, mes)} PLANILLA {planilla.group(1)}"
        self._cargar_administradoras()

        filas, sin_cuenta, total = [], [], 0.0
        for linea in texto.splitlines():
            m = PATRON_FILA.match(linea.strip())
            if not m:
                continue
            nit = m["nit"] or ""
            adm = (m["adm"] or "").strip()
            valor, mora = a_float(m["val"]), a_float(m["mora"])
            tercero = nit or self.NIT_ARUS
            nombre = adm or self.NIT_ARUS
            cuenta = self._cuenta(m["cod"], nit, adm)
            if not cuenta and valor:
                sin_cuenta.append(adm or nit or m["cod"])
            if valor:
                filas.append(self._fila(cuenta, tercero, nombre, debito=valor))
            if mora:
                filas.append(self._fila(self.CODPUC_INTERESES_MORA, tercero, nombre, debito=mora))
            total += valor + mora

        if not filas:
            raise ValueError("No se encontraron administradoras en la planilla.")

        total_pdf = PATRON_TOTAL.search(texto)
        if total_pdf and abs(a_float(total_pdf.group(1)) - total) > TOLERANCIA:
            raise ValueError(
                f"La suma de administradoras ({total:,.0f}) no coincide con el Total a Pagar "
                f"({a_float(total_pdf.group(1)):,.0f})."
            )

        self._advertir_sin_cuenta(sin_cuenta)
        filas.append(self._contrapartida(round(self._suma(filas, "debito"), 2)))
        return self._comprobante(filas, ano, mes, dia, detalle)

    def calcular(self, lin, nit_empresa="", primer_consecutivo=1):
        df = lin.copy()
        for col in ["codpuc_auto", "tercero_ref", "codpuc_manual", "tercero_manual", "docrela", "nombre"]:
            df[col] = df[col].fillna("").astype(str)
        df["tipo"] = df["tipo"].fillna("").astype(str).str.strip().str.upper()
        df["codpuc"] = df["codpuc_manual"].where(df["codpuc_manual"] != "", df["codpuc_auto"])
        df["tercero"] = df["tercero_manual"].where(df["tercero_manual"] != "", df["tercero_ref"])
        df["numero"] = self._numerar(df["mov"], df["tipo"], primer_consecutivo)
        return df


class GeneradorPlanillasMultiples(GeneradorSeguridadSocial):
    def __init__(self, generadores):
        GeneradorArchivoPlano.__init__(self, generadores[0].ruta_origen)
        self.generadores = generadores
        self.nit_empresa = ""
        self.advertencias = []

    def preparar(self):
        partes, nits, self.advertencias = [], set(), []
        for mov, gen in enumerate(self.generadores):
            nombre = os.path.basename(gen.ruta_origen)
            try:
                lin = gen.preparar()
            except Exception as e:
                raise ValueError(f"{nombre}:\n{e}") from e
            partes.append(lin.assign(mov=mov))
            self.advertencias += [f"{nombre}:\n{a}" for a in gen.advertencias]
            if gen.nit_empresa:
                nits.add(gen.nit_empresa)

        if len(nits) > 1:
            raise ValueError(f"Las planillas pertenecen a empresas diferentes (NIT {', '.join(sorted(nits))}).")
        detalles = pd.concat(partes).drop_duplicates("mov")["detalle"]
        repetidas = detalles[detalles.duplicated()].unique()
        if len(repetidas):
            raise ValueError("Planilla cargada más de una vez:\n- " + "\n- ".join(repetidas))

        self.nit_empresa = next(iter(nits), "")
        return pd.concat(partes, ignore_index=True)


Generador = GeneradorSeguridadSocial


def generador_para(clase, rutas, **opciones):
    generadores = [clase(r, **opciones) for r in rutas]
    return generadores[0] if len(generadores) == 1 else GeneradorPlanillasMultiples(generadores)