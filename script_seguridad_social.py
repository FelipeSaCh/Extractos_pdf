#script_seguridad_social.py
import json
import re
import unicodedata

import numpy as np
import pandas as pd
import pdfplumber

from archivo_plano import GeneradorArchivoPlano

RUTA_ADMINISTRADORAS = r"\\SRV-CONTAB\Servidor de Archivos\2025 SERVIDOR\APLICATIVO FELIPE\Documentos_auditor\administradoras.json"

PATRON_FILA = re.compile(
    r"^(?:N(?P<nit>\d{6,15})\s+)?(?P<cod>\S+)(?:\s+(?P<adm>.*?))?\s+(?P<afi>\d+)"
    r"\s+\$\s*(?P<val>[\d.,]+)\s+\$\s*(?P<mora>[\d.,]+)\s*$"
)
PATRON_FECHA = re.compile(r"Fecha\s*:?\s*(\d{4})-(\d{1,2})-(\d{1,2})")
PATRON_PLANILLA = re.compile(r"N[úu]mero\s+Planilla\s*:?\s*(\d+)", re.I)
PATRON_NIT_EMPRESA = re.compile(r"^NIT\s+[A-Z]{1,2}\s*(\d{5,15})", re.M)
PATRON_TOTAL = re.compile(r"Total\s+a\s+Pagar\s*:?\s*\$\s*([\d.,]+)", re.I)


def _monto(txt):
    txt = str(txt or "").strip()
    txt = txt.replace(".", "").replace(",", ".") if "," in txt else txt.replace(".", "")
    return float(txt or 0)


def _normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", texto).strip()


def _digitos(valor):
    if isinstance(valor, float):
        valor = "" if np.isnan(valor) else int(valor)
    return re.sub(r"\D", "", str(valor if valor is not None else ""))


def _cargar_administradoras(ruta):
    with open(ruta, encoding="utf-8-sig") as f:
        data = json.load(f)
    if isinstance(data, dict):
        listas = [v for v in data.values() if isinstance(v, list) and v and isinstance(v[0], dict)]
        if listas:
            data = listas[0]
    df = pd.DataFrame(data)
    cols = {_normalizar(c): c for c in df.columns}

    def col(clave, obligatoria=True):
        c = next((o for n, o in cols.items() if clave in n), None)
        if c is None:
            if obligatoria:
                raise ValueError(f"administradoras.json no contiene la columna '{clave}'.")
            return pd.Series("", index=df.index)
        return df[c]

    return pd.DataFrame({
        "nombre": col("NOMBRE").map(_normalizar),
        "nit": col("NIT").map(_digitos),
        "categoria": col("CATEGOR", False).map(_normalizar),
        "cuenta": col("CUENTA").map(_digitos),
    })


class GeneradorSeguridadSocial(GeneradorArchivoPlano):
    CLASE_SS = "SEGURIDAD_SOCIAL"
    COMPROBANTES_PREDETERMINADOS = {CLASE_SS: "PS1"}
    CLAVE_CONSECUTIVO = "consecutivo_seguridad_social"
    NIT_ARUS = "900089104"
    NOMBRE_OPERADOR = "ENLACE OPERATIVO"
    CODPUC_INTERESES_MORA = "530520001"

    def __init__(self, ruta_pdf, ruta_administradoras=RUTA_ADMINISTRADORAS):
        super().__init__(ruta_pdf)
        self.ruta_administradoras = ruta_administradoras
        self.nit_empresa = ""
        self.advertencias = []
        self._adm = None

    def _leer_pdf(self):
        with pdfplumber.open(self.ruta_origen) as pdf:
            return "\n".join(p.extract_text() or "" for p in pdf.pages)

    def _buscar_cuenta(self, nit, nombre):
        adm = self._adm[self._adm["cuenta"] != ""]
        cand = adm[adm["nit"] == nit] if nit else adm.iloc[0:0]
        if cand["cuenta"].nunique() > 1:
            cand = cand[~cand["categoria"].str.contains("CESANT", na=False)]
        if cand["cuenta"].nunique() == 1:
            return cand["cuenta"].iloc[0]
        nom = _normalizar(nombre)
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

    @staticmethod
    def _fila(cuenta, tercero, nombre, debito=np.nan, credito=np.nan):
        return {"codpuc_auto": cuenta, "tercero_ref": tercero, "nombre": nombre,
                "debito": debito, "credito": credito}

    def preparar(self):
        texto = self._leer_pdf()
        fecha = PATRON_FECHA.search(texto)
        planilla = PATRON_PLANILLA.search(texto)
        if not fecha or not planilla:
            raise ValueError("No se encontró la Fecha o el Número de Planilla. Verifique que sea una planilla de Arus (Enlace Operativo).")

        ano, mes, dia = map(int, fecha.groups())
        nit_emp = PATRON_NIT_EMPRESA.search(texto)
        self.nit_empresa = nit_emp.group(1) if nit_emp else ""
        detalle = f"PAGO SEGURIDAD SOCIAL {self.MESES[mes - 1]} PLANILLA {planilla.group(1)}"
        self._adm = _cargar_administradoras(self.ruta_administradoras)

        filas, sin_cuenta, total = [], [], 0.0
        for linea in texto.splitlines():
            m = PATRON_FILA.match(linea.strip())
            if not m:
                continue
            nit = m["nit"] or ""
            adm = (m["adm"] or "").strip()
            valor, mora = _monto(m["val"]), _monto(m["mora"])
            tercero = nit or self.NIT_ARUS
            nombre = adm or self.NIT_ARUS
            cuenta = self._buscar_cuenta(nit, adm)
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
        if total_pdf and abs(_monto(total_pdf.group(1)) - total) > 1:
            raise ValueError(
                f"La suma de administradoras ({total:,.0f}) no coincide con el Total a Pagar "
                f"({_monto(total_pdf.group(1)):,.0f})."
            )

        self.advertencias = (
            [f"Sin cuenta contable en administradoras.json:\n- " + "\n- ".join(sin_cuenta)] if sin_cuenta else []
        )

        return pd.DataFrame(filas).assign(
            mov=0, clase=self.CLASE_SS, tipo=self.COMPROBANTES_PREDETERMINADOS[self.CLASE_SS],
            ano=ano, mes=mes, dia=dia, detalle=detalle,
            codpuc_manual="", tercero_manual="", docrela="",
        )

    def calcular(self, lin, nit_empresa="", primer_consecutivo=1):
        df = lin.copy()
        for col in ["codpuc_auto", "tercero_ref", "codpuc_manual", "tercero_manual", "docrela", "nombre"]:
            df[col] = df[col].fillna("").astype(str)
        df["tipo"] = df["tipo"].fillna("").astype(str).str.strip().str.upper()
        df["codpuc"] = df["codpuc_manual"].where(df["codpuc_manual"] != "", df["codpuc_auto"])
        df["tercero"] = df["tercero_manual"].where(df["tercero_manual"] != "", df["tercero_ref"])
        inicio = self._a_entero(primer_consecutivo)
        df["numero"] = [f"{n:010d}" for n in pd.factorize(df["mov"])[0] + inicio]
        return df