"""Archivo plano contable (26 columnas, Excel 5.0/95) a partir de la hoja Datos del extracto."""
import datetime
import os
import re

import numpy as np
import pandas as pd

from infraestructura.excel import convertir_a_xls95, leer_cuenta_bancaria
from nucleo.constantes import HOJA_DATOS, MESES
from nucleo.fechas import partes_fecha
from nucleo.numeros import a_entero, serie_a_numero
from nucleo.validaciones import PATRON_TERCERO_REFERENCIA


class GeneradorArchivoPlano:
    COLUMNAS_PLANO = [
        "libro", "tipo", "numero", "ano", "mes", "dia", "detalle", "anulado",
        "codpuc", "tercero", "ccosto", "sccosto", "docrela", "cuota", "artrela",
        "artcant", "debito", "credito", "nombre", "base", "base2", "fechaven",
        "area", "numexterno", "anexo", "bodega"
    ]

    ALIAS_COLUMNAS = {
        "FECHA": ("FECHA", "FECHA DEL MOVIMIENTO", "FECHA MOVIMIENTO", "FECHA TRANSACCION", "FECHA TRANSACCIÓN"),
        "DESCRIPCIÓN": ("DESCRIPCIÓN", "DESCRIPCION", "CLASE DE MOVIMIENTO", "CONCEPTO", "DETALLE",
                        "TRANSACCION", "TRANSACCIÓN"),
        "VALOR": ("VALOR", "MONTO", "IMPORTE", "VALOR TRANSACCION", "VALOR TRANSACCIÓN"),
    }
    COLUMNAS_DEBITO = ("DEBITOS", "DÉBITOS", "DEBITO", "DÉBITO", "CARGOS", "CARGO", "VALOR DEBITO", "VALOR DÉBITO")
    COLUMNAS_CREDITO = ("CREDITOS", "CRÉDITOS", "CREDITO", "CRÉDITO", "ABONOS", "ABONO", "VALOR CREDITO", "VALOR CRÉDITO")
    COLUMNAS_REFERENCIA = ("REFERENCIA 1", "REFERENCIA", "DCTO.", "DOCUMENTO")
    COLUMNAS_TERCERO = ("NIT", "CEDULA", "CÉDULA", "REFERENCIA 1", "REFERENCIA", "DCTO.")

    CONCEPTOS_BANCARIOS = [
        ("GMF 4X1000", r"4\s*X\s*1\.?000|GRAVAMEN|\bGMF\b"),
        ("RETENCION EN LA FUENTE", r"RET(?:ENCI[OÓ]N|E)\.?\s*(?:EN\s+LA\s+)?FUENTE|\bRTE\.?\s*FTE\b"),
        ("IVA", r"\bIVA\b"),
        ("CUOTA PLAN CANAL NEGOCIOS", r"CUOTA\s+PLAN\s+CANAL\s+NEGOCIOS"),
        ("CUOTA DE MANEJO", r"CUOTA\s+(?:DE\s+)?MANEJO"),
        ("TARJETA DE CREDITO", r"TARJETA\s+DE\s+CR[EÉ]DITO"),
        ("COMISION", r"COMISI[OÓ]N"),
        ("SERVICIO", r"SERVICIO"),
        ("TARJ DEB", r"TARJ\s+DEB"),
        ("INTERESES AHORROS", r"INTERESES\s+AHORROS|ABONO\s+INTERESES|\bRENDIMIENTOS?\b"),
    ]
    CODPUC_BANCARIOS = {
        "GMF 4X1000": "530595001",
        "RETENCION EN LA FUENTE": "135515001",
        "IVA": "530505002",
        "CUOTA PLAN CANAL NEGOCIOS": "530505001",
        "CUOTA DE MANEJO": "530505001",
        "TARJETA DE CREDITO": "530505003",
        "COMISION": "530515001",
        "SERVICIO": "530505001",
        "TARJ DEB": "530505001",
        "INTERESES AHORROS": "421005001",
    }
    CODPUC_BANCO = "112005001"
    CODPUC_PROVEEDORES = "220501001"
    CODPUC_CLIENTES = "130505001"
    CODPUC_NOMINA = "250505001"
    CODPUC_GASTO_BANCARIO = "530505001"
    CODPUC_RETIROS = "110505001"
    CLAVE_CONSECUTIVO = "consecutivo"
    PATRON_RETIROS = r"\bRETIROS?\b"

    CLASE_EGRESO = "EGRESO"
    CLASE_NOMINA = "NOMINA"
    CLASE_RECIBO = "RECIBO"
    CLASE_BANCARIO = "BANCARIO"

    COMPROBANTES_PREDETERMINADOS = {
        CLASE_EGRESO: "EG1",
        CLASE_NOMINA: "EG2",
        CLASE_RECIBO: "RC1",
        CLASE_BANCARIO: "AJB",
    }

    DETALLE_CONTRAPARTIDA = {
        CLASE_EGRESO: "TOTAL EGRESOS",
        CLASE_NOMINA: "TOTAL PAGOS DE NOMINA",
        CLASE_RECIBO: "TOTAL RECAUDOS",
        CLASE_BANCARIO: "TOTAL MOVIMIENTOS BANCARIOS",
    }
    PREFIJOS_BANCARIO = ("AJ", "CC")
    ORDEN_PREFIJOS = {"AJ": 0, "CC": 0, "EG": 1, "RC": 2}

    LADO_BANCO = "BANCO"
    LADO_CONTRA = "CONTRA"

    PATRON_NOMINA = r"PAGO\s+(?:A\s+|DE\s+)?N[OÓ]MIN\w*"
    DETALLE_NOMINA = "PAGO A NOMIN"

    MESES = MESES

    def __init__(self, ruta_excel_origen):
        self.ruta_origen = ruta_excel_origen
        self.cuenta_bancaria = ""

    # ------------------------------------------------------------------ utilidades
    @staticmethod
    def _columna(df, candidatos):
        return next((c for c in candidatos if c in df.columns), None)

    @staticmethod
    def _texto(serie):
        return serie.fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

    @staticmethod
    def _tercero_valido(texto):
        texto = str(texto or "").strip()
        return texto if PATRON_TERCERO_REFERENCIA.match(texto) else ""

    @staticmethod
    def _entero_o_vacio(serie):
        return serie.map(lambda x: str(int(x)) if pd.notna(x) else "")

    @staticmethod
    def _fin_de_mes(ano, mes):
        return pd.to_datetime(pd.DataFrame({"year": ano.astype(int), "month": mes.astype(int), "day": 1})) \
            + pd.offsets.MonthEnd(0)

    @classmethod
    def _periodo(cls, mes, ano):
        return mes.astype(int).map(lambda m: cls.MESES[m - 1]) + " " + ano.astype(int).astype(str)

    @staticmethod
    def _numerar(clave, tipo, primer_consecutivo):
        """Número de 10 dígitos por comprobante (clave), consecutivo dentro de cada tipo."""
        comps = pd.DataFrame({"clave": clave.values, "tipo": tipo.values}).drop_duplicates("clave")
        if isinstance(primer_consecutivo, dict):
            inicio = comps["tipo"].map(lambda t: a_entero(primer_consecutivo.get(t, 1)))
        else:
            inicio = a_entero(primer_consecutivo)
        numeros = inicio + comps.groupby("tipo").cumcount()
        return clave.map(dict(zip(comps["clave"], numeros))).map(lambda n: f"{int(n):010d}")

    # ------------------------------------------------------------------ lectura
    def _normalizar_columnas(self, df):
        df = df.rename(columns=lambda c: re.sub(r"\s+", " ", str(c)).strip().upper())
        for canon, alias in self.ALIAS_COLUMNAS.items():
            col = None if canon in df.columns else self._columna(df, alias)
            if col:
                df = df.rename(columns={col: canon})
        if "VALOR" not in df.columns:
            deb = self._columna(df, self.COLUMNAS_DEBITO)
            cre = self._columna(df, self.COLUMNAS_CREDITO)
            if deb or cre:
                abonos = serie_a_numero(df[cre]).fillna(0.0).abs() if cre else 0.0
                cargos = serie_a_numero(df[deb]).fillna(0.0).abs() if deb else 0.0
                df["VALOR"] = abonos - cargos
        return df

    def _extraer_fechas(self, df):
        if "FECHA" not in df.columns:
            vacio = pd.Series(np.nan, index=df.index)
            return vacio, vacio.copy(), vacio.copy()
        partes = pd.DataFrame(df["FECHA"].map(partes_fecha).tolist(),
                              columns=["ano", "mes", "dia"], index=df.index, dtype=float)
        valida = partes["mes"].between(1, 12) & partes["dia"].between(1, 31)
        partes = partes.where(valida)
        ano = partes["ano"].where(partes["ano"] >= 2000)
        hoy = datetime.date.today()
        ano_inferido = pd.Series(np.where(partes["mes"] > hoy.month, hoy.year - 1, hoy.year), index=df.index)
        partes["ano"] = ano.mask(valida & ano.isna(), ano_inferido)
        return partes["ano"], partes["mes"], partes["dia"]

    def _construir_movimientos(self, df):
        valores = serie_a_numero(df["VALOR"]).fillna(0.0) if "VALOR" in df.columns else pd.Series(0.0, index=df.index)
        ano, mes, dia = self._extraer_fechas(df)
        desc = df.get("DESCRIPCIÓN", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
        col_ref = self._columna(df, self.COLUMNAS_REFERENCIA)
        ref = self._texto(df[col_ref]) if col_ref else pd.Series("", index=df.index)

        tercero = pd.Series("", index=df.index)
        for col in self.COLUMNAS_TERCERO:
            if col in df.columns:
                tercero = tercero.mask(tercero == "", self._texto(df[col]).map(self._tercero_valido))

        mov = pd.DataFrame({
            "clase": "", "ano": ano, "mes": mes, "dia": dia, "desc": desc,
            "detalle": (desc + " " + ref).str.strip(), "tercero": tercero, "valor": valores, "codpuc_contra": "",
        })
        return mov[mov["valor"] != 0].reset_index(drop=True)

    # ------------------------------------------------------------------ clasificación
    def _consolidar_bancarios(self, mov):
        desc_upper = mov["desc"].str.upper()
        concepto = pd.Series("", index=mov.index)
        for nombre, patron in self.CONCEPTOS_BANCARIOS:
            concepto = concepto.mask((concepto == "") & desc_upper.str.contains(patron, regex=True, na=False), nombre)

        mask = (concepto != "") & mov["mes"].between(1, 12)
        if not mask.any():
            return mov

        grp = pd.DataFrame({
            "ano": mov["ano"].fillna(datetime.date.today().year),
            "mes": mov["mes"],
            "concepto": concepto,
            "valor": mov["valor"],
        })[mask].groupby(["ano", "mes", "concepto"], as_index=False, sort=True)["valor"].sum()
        grp = grp[grp["valor"] != 0].reset_index(drop=True)
        if grp.empty:
            return mov[~mask].reset_index(drop=True)

        fin_mes = self._fin_de_mes(grp["ano"], grp["mes"])
        cuenta = f" CUENTA {self.cuenta_bancaria}" if self.cuenta_bancaria else ""
        consolidado = pd.DataFrame({
            "clase": self.CLASE_BANCARIO,
            "ano": fin_mes.dt.year,
            "mes": fin_mes.dt.month,
            "dia": fin_mes.dt.day,
            "desc": "",
            "detalle": grp["concepto"] + " " + self._periodo(fin_mes.dt.month, fin_mes.dt.year) + cuenta,
            "tercero": "",
            "valor": grp["valor"],
            "codpuc_contra": grp["concepto"].map(self.CODPUC_BANCARIOS).fillna(""),
        })
        return pd.concat([mov[~mask], consolidado], ignore_index=True)

    def _agrupar_nomina(self, mov):
        mask = (mov["clase"] == "") & mov["detalle"].str.contains(self.PATRON_NOMINA, case=False, regex=True, na=False)
        if not mask.any():
            return mov
        benef = mov.loc[mask, "detalle"].str.extract(
            rf"(?i){self.PATRON_NOMINA}[\s:\-]*(.*)$", expand=False).fillna("").str.strip()
        agrupado = (
            mov[mask].assign(_benef=benef)
            .groupby(["ano", "mes", "dia"], as_index=False, dropna=False)
            .agg(valor=("valor", "sum"), _benef=("_benef", lambda s: ", ".join(dict.fromkeys(v for v in s if v))))
        )
        agrupado = agrupado.assign(
            clase=self.CLASE_NOMINA, desc="", detalle=(self.DETALLE_NOMINA + " " + agrupado["_benef"]).str.strip(),
            tercero="", codpuc_contra="",
        ).drop(columns="_benef")
        return pd.concat([mov[~mask], agrupado[agrupado["valor"] != 0]], ignore_index=True)

    def _asignar_clases(self, mov):
        pendiente = mov["clase"] == ""
        mov.loc[pendiente, "clase"] = np.where(mov.loc[pendiente, "valor"] < 0, self.CLASE_EGRESO, self.CLASE_RECIBO)
        return mov

    def preparar(self):
        try:
            with pd.ExcelFile(self.ruta_origen) as xl:
                df = pd.read_excel(xl, sheet_name=HOJA_DATOS)
                self.cuenta_bancaria = leer_cuenta_bancaria(xl)
        except ValueError:
            raise ValueError("El archivo seleccionado no contiene la hoja 'Datos'.")
        df = self._normalizar_columnas(df)
        if "VALOR" not in df.columns:
            raise ValueError("No se identificó la columna de valor (VALOR o DÉBITOS/CRÉDITOS) en la hoja 'Datos'.")
        if "FECHA" in df.columns:
            df = df[df["FECHA"].notna() & df["FECHA"].astype(str).str.strip().ne("")].reset_index(drop=True)

        mov = self._asignar_clases(self._agrupar_nomina(self._consolidar_bancarios(self._construir_movimientos(df))))
        mov = mov.drop(columns="desc").rename(columns={"tercero": "tercero_ref"}).reset_index(drop=True)
        mov["tipo"] = mov["clase"].map(self.COMPROBANTES_PREDETERMINADOS)
        lin = (
            pd.concat([mov.assign(lado=self.LADO_BANCO), mov.assign(lado=self.LADO_CONTRA)])
            .rename_axis("mov")
            .reset_index()
        )
        return lin.assign(codpuc_manual="", tercero_manual="", docrela="", archivo=0)

    # ------------------------------------------------------------------ comprobantes
    def clave_comprobante(self, df):
        tipo = df["tipo"].fillna("").astype(str).str.strip().str.upper()
        archivo = df["archivo"].astype(str) if "archivo" in df.columns else "0"
        periodo = df["ano"].astype(str) + "|" + df["mes"].astype(str)
        grupo = periodo.where(tipo.str[:2].isin(self.PREFIJOS_BANCARIO), "M" + df["mov"].astype(str))
        return tipo + "|" + archivo + "|" + grupo

    def conteo_comprobantes(self, lin):
        return lin.assign(_clave=self.clave_comprobante(lin)).drop_duplicates("_clave")["tipo"].value_counts()

    def _consolidar_banco(self, banco):
        mensual = banco["tipo"].str[:2].isin(self.PREFIJOS_BANCARIO)
        individuales = banco[~mensual].assign(tercero_ref="")
        banco = banco[mensual]
        if banco.empty:
            return individuales
        primero = lambda s: next((v for v in s if v), "")  # noqa: E731
        agg = (
            banco.assign(_clave=self.clave_comprobante(banco), _idx=banco.index)
            .groupby("_clave", sort=False)
            .agg(_idx=("_idx", "first"), mov=("mov", "first"), tipo=("tipo", "first"), clase=("clase", "first"),
                 lado=("lado", "first"), archivo=("archivo", "first"), ano=("ano", "first"), mes=("mes", "first"),
                 valor=("valor", "sum"), codpuc_manual=("codpuc_manual", primero),
                 tercero_manual=("tercero_manual", primero), docrela=("docrela", primero))
        )
        agg = agg[agg["valor"].round(2) != 0]
        if agg.empty:
            return individuales

        valida = agg["mes"].between(1, 12) & agg["ano"].notna()
        mes, ano = agg["mes"].where(valida, 1), agg["ano"].where(valida, 2000)
        agg["dia"] = self._fin_de_mes(ano, mes).dt.day.where(valida)
        agg["detalle"] = (agg["clase"].map(self.DETALLE_CONTRAPARTIDA).fillna("TOTAL " + agg["tipo"])
                          + " " + self._periodo(mes, ano).where(valida, "")).str.strip()
        agg["tercero_ref"] = ""
        agg["codpuc_contra"] = ""
        return pd.concat([individuales, agg.set_index("_idx").rename_axis(None)])

    def calcular(self, lin, nit_empresa="", primer_consecutivo=1):
        df = lin.copy()
        if "archivo" not in df.columns:
            df["archivo"] = 0
        for col in ["codpuc_contra", "tercero_ref", "codpuc_manual", "tercero_manual", "docrela"]:
            df[col] = df[col].fillna("").astype(str)
        df["tipo"] = df["tipo"].fillna("").astype(str).str.strip().str.upper()

        es_banco = df["lado"] == self.LADO_BANCO
        df = pd.concat([df[~es_banco], self._consolidar_banco(df[es_banco])])

        prefijo = df["tipo"].str[:2]
        mensual = prefijo.isin(self.PREFIJOS_BANCARIO)
        entrada = df["valor"] > 0
        es_banco = df["lado"] == self.LADO_BANCO
        en_debito = es_banco == entrada

        contra = np.select(
            [
                mensual,
                (prefijo == "EG") & (df["clase"] == self.CLASE_NOMINA),
                prefijo == "EG",
                prefijo == "RC",
                entrada,
            ],
            [
                df["codpuc_contra"].replace("", self.CODPUC_GASTO_BANCARIO),
                self.CODPUC_NOMINA,
                self.CODPUC_PROVEEDORES,
                self.CODPUC_CLIENTES,
                self.CODPUC_CLIENTES,
            ],
            default=self.CODPUC_PROVEEDORES,
        )
        automatico = pd.Series(np.where(es_banco, self.CODPUC_BANCO, contra), index=df.index)
        df["codpuc"] = df["codpuc_manual"].where(df["codpuc_manual"] != "", automatico)

        nit = str(nit_empresa or "").strip()
        tercero = df["tercero_ref"].where(df["tercero_ref"] != "", nit)
        df["tercero"] = df["tercero_manual"].where(df["tercero_manual"] != "", tercero)

        importe = df["valor"].abs()
        df["debito"] = importe.where(en_debito)
        df["credito"] = importe.where(~en_debito)

        df = (
            df.assign(_rango=prefijo.map(self.ORDEN_PREFIJOS).fillna(len(self.ORDEN_PREFIJOS)),
                      _banco=es_banco, _credito=~en_debito,
                      _dia=df["dia"].where(~mensual, 0), _mov=df["mov"].where(~mensual, -1))
            .sort_values(["archivo", "_rango", "tipo", "ano", "mes", "_dia", "_mov", "_banco", "dia", "detalle",
                          "mov", "_credito"], kind="stable", na_position="last")
            .drop(columns=["_rango", "_banco", "_credito", "_dia", "_mov"])
        )
        df["numero"] = self._numerar(self.clave_comprobante(df), df["tipo"], primer_consecutivo)
        return df

    # ------------------------------------------------------------------ salida
    def generar(self, lin=None, nit_empresa="", primer_consecutivo=1, ruta_destino_final=None):
        if not ruta_destino_final:
            ruta_destino_final = f"{os.path.splitext(self.ruta_origen)[0]}_PLANO.xls"

        lin = self.preparar() if lin is None else lin
        if lin.empty:
            return ruta_destino_final

        df = self.calcular(lin, nit_empresa, primer_consecutivo)
        if df["tipo"].eq("").any():
            raise ValueError("Existen movimientos sin comprobante asignado.")

        vacio_si_nan = lambda s: s.astype(object).where(s.notna(), "")  # noqa: E731
        df_plano = pd.DataFrame({
            "tipo": df["tipo"],
            "numero": df["numero"],
            "ano": self._entero_o_vacio(df["ano"]),
            "mes": self._entero_o_vacio(df["mes"]),
            "nombre": df["nombre"] if "nombre" in df.columns else "",
            "dia": self._entero_o_vacio(df["dia"]),
            "detalle": df["detalle"],
            "codpuc": df["codpuc"],
            "tercero": df["tercero"],
            "docrela": df["docrela"],
            "debito": vacio_si_nan(df["debito"]),
            "credito": vacio_si_nan(df["credito"]),
        }).reindex(columns=self.COLUMNAS_PLANO, fill_value="")

        ruta_temp_xlsx = os.path.splitext(ruta_destino_final)[0] + ".xlsx"
        df_plano.to_excel(ruta_temp_xlsx, index=False, engine="openpyxl")
        convertir_a_xls95(ruta_temp_xlsx, ruta_destino_final)
        return ruta_destino_final


class GeneradorExtractosMultiples(GeneradorArchivoPlano):
    def __init__(self, rutas):
        super().__init__(rutas[0])
        self.generadores = [GeneradorArchivoPlano(r) for r in rutas]

    def preparar(self):
        partes, base = [], 0
        for archivo, gen in enumerate(self.generadores):
            try:
                lin = gen.preparar()
            except Exception as e:
                raise ValueError(f"{os.path.basename(gen.ruta_origen)}:\n{e}") from e
            if lin.empty:
                continue
            partes.append(lin.assign(mov=lin["mov"] + base, archivo=archivo))
            base += int(lin["mov"].max()) + 1
        return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()