#archivo_plano.py
import pandas as pd
import numpy as np
import re
import os
import datetime
import win32com.client as win32
import pythoncom


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

    MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
             "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

    def __init__(self, ruta_excel_origen):
        self.ruta_origen = ruta_excel_origen

    @staticmethod
    def _columna(df, candidatos):
        return next((c for c in candidatos if c in df.columns), None)

    @staticmethod
    def _texto(serie):
        return serie.fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()

    @staticmethod
    def _a_numero(serie):
        if not pd.api.types.is_object_dtype(serie):
            return pd.to_numeric(serie, errors='coerce')
        txt = serie.fillna("").astype(str).str.strip()
        neg = txt.str.contains(r"^\s*-|-\s*$|^\(.*\)$", regex=True)
        txt = txt.str.replace(r"[^\d.,]", "", regex=True)
        coma_dec = txt.str.contains(r",\d{1,2}$", regex=True)
        txt = txt.where(~coma_dec, txt.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        txt = txt.where(coma_dec, txt.str.replace(",", "", regex=False))
        txt = txt.where(txt.str.count(r"\.") <= 1, txt.str.replace(".", "", regex=False))
        num = pd.to_numeric(txt, errors='coerce')
        return num.where(~neg, -num)

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
                abonos = self._a_numero(df[cre]).fillna(0.0).abs() if cre else 0.0
                cargos = self._a_numero(df[deb]).fillna(0.0).abs() if deb else 0.0
                df["VALOR"] = abonos - cargos
        return df

    def _limpiar_y_convertir_valores(self, df):
        if 'VALOR' not in df.columns:
            return pd.Series(0.0, index=df.index)
        return self._a_numero(df['VALOR']).fillna(0.0)

    def _es_tercero_valido(self, texto):
        if pd.isna(texto) or not texto:
            return ""
        texto_limpio = str(texto).strip()
        if re.match(r"^\d{6,15}(?:-\d)?$", texto_limpio):
            return texto_limpio
        return ""

    @staticmethod
    def _entero_o_vacio(serie):
        return serie.map(lambda x: str(int(x)) if pd.notna(x) else "")

    @staticmethod
    def _a_entero(valor):
        try:
            return max(int(str(valor).strip()), 1)
        except (TypeError, ValueError):
            return 1

    def _extraer_fechas(self, df):
        if 'FECHA' not in df.columns:
            vacio = pd.Series(np.nan, index=df.index)
            return vacio, vacio.copy(), vacio.copy()

        partes = pd.DataFrame(df['FECHA'].map(self._partes_fecha).tolist(),
                              columns=['ano', 'mes', 'dia'], index=df.index, dtype=float)
        valida = partes['mes'].between(1, 12) & partes['dia'].between(1, 31)
        partes = partes.where(valida)
        ano = partes['ano'].where(partes['ano'] >= 2000)
        hoy = datetime.date.today()
        ano_inferido = pd.Series(np.where(partes['mes'] > hoy.month, hoy.year - 1, hoy.year), index=df.index)
        partes['ano'] = ano.mask(valida & ano.isna(), ano_inferido)
        return partes['ano'], partes['mes'], partes['dia']

    @staticmethod
    def _partes_fecha(valor):
        if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
            return np.nan, np.nan, np.nan
        if isinstance(valor, (datetime.date, pd.Timestamp)):
            return valor.year, valor.month, valor.day
        texto = f"{valor:.2f}" if isinstance(valor, float) else str(valor).strip()
        m = re.match(r'^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})', texto)
        if m:
            return int(m[1]), int(m[2]), int(m[3])
        m = re.match(r'^(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2}|\d{4}))?(?!\d)', texto)
        if m:
            ano = int(m[3]) if m[3] else np.nan
            return (ano + 2000 if ano < 100 else ano) if m[3] else np.nan, int(m[2]), int(m[1])
        return np.nan, np.nan, np.nan

    def _construir_movimientos(self, df):
        n = len(df)
        valores = self._limpiar_y_convertir_valores(df)
        ano, mes, dia = self._extraer_fechas(df)

        desc = df.get("DESCRIPCIÓN", pd.Series([""] * n, index=df.index))
        desc = desc.fillna("").astype(str).str.strip()

        col_ref = self._columna(df, self.COLUMNAS_REFERENCIA)
        ref = self._texto(df[col_ref]) if col_ref else pd.Series("", index=df.index)
        detalle = (desc + " " + ref).str.strip()

        tercero = pd.Series("", index=df.index)
        for col in self.COLUMNAS_TERCERO:
            if col in df.columns:
                tercero = tercero.mask(tercero == "", self._texto(df[col]).apply(self._es_tercero_valido))

        mov = pd.DataFrame({
            'clase': "",
            'ano': ano,
            'mes': mes,
            'dia': dia,
            'desc': desc,
            'detalle': detalle,
            'tercero': tercero,
            'valor': valores,
            'codpuc_contra': "",
        })
        return mov[mov['valor'] != 0].reset_index(drop=True)

    def _consolidar_bancarios(self, mov):
        desc_upper = mov['desc'].str.upper()
        concepto = pd.Series("", index=mov.index)
        for nombre, patron in self.CONCEPTOS_BANCARIOS:
            coincide = desc_upper.str.contains(patron, regex=True, na=False)
            concepto = concepto.mask((concepto == "") & coincide, nombre)

        mask = (concepto != "") & mov['mes'].between(1, 12)
        if not mask.any():
            return mov

        grp = pd.DataFrame({
            'ano': mov['ano'].fillna(datetime.date.today().year),
            'mes': mov['mes'],
            'concepto': concepto,
            'valor': mov['valor'],
        })[mask].groupby(['ano', 'mes', 'concepto'], as_index=False, sort=True)['valor'].sum()
        grp = grp[grp['valor'] != 0].reset_index(drop=True)

        if grp.empty:
            return mov[~mask].reset_index(drop=True)

        fin_mes = pd.to_datetime(pd.DataFrame({
            'year': grp['ano'].astype(int),
            'month': grp['mes'].astype(int),
            'day': 1
        })) + pd.offsets.MonthEnd(0)

        consolidado = pd.DataFrame({
            'clase': self.CLASE_BANCARIO,
            'ano': fin_mes.dt.year,
            'mes': fin_mes.dt.month,
            'dia': fin_mes.dt.day,
            'desc': "",
            'detalle': grp['concepto'] + " " + fin_mes.dt.month.map(lambda m: self.MESES[m - 1]) + " " + fin_mes.dt.year.astype(str),
            'tercero': "",
            'valor': grp['valor'],
            'codpuc_contra': grp['concepto'].map(self.CODPUC_BANCARIOS).fillna(""),
        })

        return pd.concat([mov[~mask], consolidado], ignore_index=True)

    def _agrupar_nomina(self, mov):
        mask = (mov['clase'] == "") & mov['detalle'].str.contains(self.PATRON_NOMINA, case=False, regex=True, na=False)
        if not mask.any():
            return mov

        benef = mov.loc[mask, 'detalle'].str.extract(
            rf"(?i){self.PATRON_NOMINA}[\s:\-]*(.*)$", expand=False).fillna("").str.strip()
        agrupado = (
            mov[mask].assign(_benef=benef)
            .groupby(['ano', 'mes', 'dia'], as_index=False, dropna=False)
            .agg(valor=('valor', 'sum'),
                 _benef=('_benef', lambda s: ", ".join(dict.fromkeys(v for v in s if v))))
        )
        agrupado = agrupado.assign(
            clase=self.CLASE_NOMINA,
            desc="",
            detalle=(self.DETALLE_NOMINA + " " + agrupado['_benef']).str.strip(),
            tercero="",
            codpuc_contra="",
        ).drop(columns='_benef')
        agrupado = agrupado[agrupado['valor'] != 0]

        return pd.concat([mov[~mask], agrupado], ignore_index=True)

    def _asignar_clases(self, mov):
        pendiente = mov['clase'] == ""
        mov.loc[pendiente, 'clase'] = np.where(
            mov.loc[pendiente, 'valor'] < 0, self.CLASE_EGRESO, self.CLASE_RECIBO
        )
        return mov

    def _convertir_a_excel_95(self, ruta_xlsx, ruta_xls):
        try:
            pythoncom.CoInitialize()
            excel = win32.DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False

            ruta_abs_xlsx = os.path.abspath(ruta_xlsx)
            ruta_abs_xls = os.path.abspath(ruta_xls)

            wb = excel.Workbooks.Open(ruta_abs_xlsx)
            wb.SaveAs(ruta_abs_xls, FileFormat=39)
            wb.Close(False)
            excel.Quit()

            if os.path.exists(ruta_abs_xlsx):
                os.remove(ruta_abs_xlsx)

        except Exception as e:
            raise RuntimeError(f"Fallo en la conversión a formato 5.0/95: {e}")
        finally:
            pythoncom.CoUninitialize()

    def preparar(self):
        try:
            df = pd.read_excel(self.ruta_origen, sheet_name="Datos")
        except ValueError:
            raise ValueError("El archivo seleccionado no contiene la hoja 'Datos'.")
        df = self._normalizar_columnas(df)
        if "VALOR" not in df.columns:
            raise ValueError("No se identificó la columna de valor (VALOR o DÉBITOS/CRÉDITOS) en la hoja 'Datos'.")
        if "FECHA" in df.columns:
            df = df[df["FECHA"].notna() & df["FECHA"].astype(str).str.strip().ne("")].reset_index(drop=True)

        mov = self._construir_movimientos(df)
        mov = self._consolidar_bancarios(mov)
        mov = self._agrupar_nomina(mov)
        mov = self._asignar_clases(mov)
        mov = mov.drop(columns='desc').rename(columns={'tercero': 'tercero_ref'}).reset_index(drop=True)
        mov['tipo'] = mov['clase'].map(self.COMPROBANTES_PREDETERMINADOS)

        lin = (
            pd.concat([mov.assign(lado=self.LADO_BANCO), mov.assign(lado=self.LADO_CONTRA)])
            .rename_axis('mov')
            .reset_index()
        )
        return lin.assign(codpuc_manual="", tercero_manual="", docrela="", archivo=0)

    def clave_comprobante(self, df):
        tipo = df['tipo'].fillna("").astype(str).str.strip().str.upper()
        archivo = df['archivo'].astype(str) if 'archivo' in df.columns else "0"
        periodo = df['ano'].astype(str) + "|" + df['mes'].astype(str)
        grupo = periodo.where(tipo.str[:2].isin(self.PREFIJOS_BANCARIO), "M" + df['mov'].astype(str))
        return tipo + "|" + archivo + "|" + grupo

    def conteo_comprobantes(self, lin):
        return lin.assign(_clave=self.clave_comprobante(lin)).drop_duplicates('_clave')['tipo'].value_counts()

    def _consolidar_banco(self, banco):
        mensual = banco['tipo'].str[:2].isin(self.PREFIJOS_BANCARIO)
        individuales = banco[~mensual].assign(tercero_ref="")
        banco = banco[mensual]
        if banco.empty:
            return individuales
        primero = lambda s: next((v for v in s if v), "")
        agg = (
            banco.assign(_clave=self.clave_comprobante(banco), _idx=banco.index)
            .groupby('_clave', sort=False)
            .agg(_idx=('_idx', 'first'), mov=('mov', 'first'), tipo=('tipo', 'first'), clase=('clase', 'first'),
                 lado=('lado', 'first'), archivo=('archivo', 'first'), ano=('ano', 'first'), mes=('mes', 'first'),
                 valor=('valor', 'sum'), codpuc_manual=('codpuc_manual', primero),
                 tercero_manual=('tercero_manual', primero), docrela=('docrela', primero))
        )
        agg = agg[agg['valor'].round(2) != 0]
        if agg.empty:
            return individuales

        valida = agg['mes'].between(1, 12) & agg['ano'].notna()
        fin_mes = pd.to_datetime(pd.DataFrame({
            'year': agg['ano'].where(valida, 2000).astype(int),
            'month': agg['mes'].where(valida, 1).astype(int),
            'day': 1,
        })) + pd.offsets.MonthEnd(0)
        periodo = (agg['mes'].where(valida, 1).astype(int).map(lambda m: self.MESES[m - 1]) + " "
                   + agg['ano'].where(valida, 0).astype(int).astype(str)).where(valida, "")
        agg['dia'] = fin_mes.dt.day.where(valida)
        agg['detalle'] = (agg['clase'].map(self.DETALLE_CONTRAPARTIDA).fillna("TOTAL " + agg['tipo'])
                          + " " + periodo).str.strip()
        agg['tercero_ref'] = ""
        agg['codpuc_contra'] = ""
        return pd.concat([individuales, agg.set_index('_idx').rename_axis(None)])

    def calcular(self, lin, nit_empresa="", primer_consecutivo=1):
        df = lin.copy()
        if 'archivo' not in df.columns:
            df['archivo'] = 0
        for col in ['codpuc_contra', 'tercero_ref', 'codpuc_manual', 'tercero_manual', 'docrela']:
            df[col] = df[col].fillna("").astype(str)
        df['tipo'] = df['tipo'].fillna("").astype(str).str.strip().str.upper()

        es_banco = df['lado'] == self.LADO_BANCO
        df = pd.concat([df[~es_banco], self._consolidar_banco(df[es_banco])])

        prefijo = df['tipo'].str[:2]
        mensual = prefijo.isin(self.PREFIJOS_BANCARIO)
        entrada = df['valor'] > 0
        es_banco = df['lado'] == self.LADO_BANCO
        en_debito = es_banco == entrada

        contra = np.select(
            [
                mensual,
                (prefijo == "EG") & (df['clase'] == self.CLASE_NOMINA),
                prefijo == "EG",
                prefijo == "RC",
                entrada,
            ],
            [
                df['codpuc_contra'].replace("", self.CODPUC_GASTO_BANCARIO),
                self.CODPUC_NOMINA,
                self.CODPUC_PROVEEDORES,
                self.CODPUC_CLIENTES,
                self.CODPUC_CLIENTES,
            ],
            default=self.CODPUC_PROVEEDORES,
        )
        automatico = pd.Series(np.where(es_banco, self.CODPUC_BANCO, contra), index=df.index)
        df['codpuc'] = df['codpuc_manual'].where(df['codpuc_manual'] != "", automatico)

        nit = str(nit_empresa or "").strip()
        tercero = df['tercero_ref'].where(df['tercero_ref'] != "", nit)
        df['tercero'] = df['tercero_manual'].where(df['tercero_manual'] != "", tercero)

        importe = df['valor'].abs()
        df['debito'] = importe.where(en_debito)
        df['credito'] = importe.where(~en_debito)

        df = (
            df.assign(_rango=prefijo.map(self.ORDEN_PREFIJOS).fillna(len(self.ORDEN_PREFIJOS)),
                      _banco=es_banco, _credito=~en_debito,
                      _dia=df['dia'].where(~mensual, 0), _mov=df['mov'].where(~mensual, -1))
            .sort_values(['archivo', '_rango', 'tipo', 'ano', 'mes', '_dia', '_mov', '_banco', 'dia', 'detalle',
                          'mov', '_credito'], kind='stable', na_position='last')
            .drop(columns=['_rango', '_banco', '_credito', '_dia', '_mov'])
        )
        clave = self.clave_comprobante(df)
        comps = df.assign(_clave=clave).drop_duplicates('_clave')[['_clave', 'tipo']]
        if isinstance(primer_consecutivo, dict):
            inicio = comps['tipo'].map(lambda t: self._a_entero(primer_consecutivo.get(t, 1)))
        else:
            inicio = pd.Series(self._a_entero(primer_consecutivo), index=comps.index)
        numeros = inicio + comps.groupby('tipo').cumcount()
        df['numero'] = clave.map(dict(zip(comps['_clave'], numeros))).map(lambda n: f"{n:010d}")
        return df

    def generar(self, lin=None, nit_empresa="", primer_consecutivo=1, ruta_destino_final=None):
        if not ruta_destino_final:
            base, _ = os.path.splitext(self.ruta_origen)
            ruta_destino_final = f"{base}_PLANO.xls"

        lin = self.preparar() if lin is None else lin
        if lin.empty:
            return ruta_destino_final

        df = self.calcular(lin, nit_empresa, primer_consecutivo)
        if df['tipo'].eq("").any():
            raise ValueError("Existen movimientos sin comprobante asignado.")

        df_plano = pd.DataFrame({
            'tipo': df['tipo'],
            'numero': df['numero'],
            'ano': self._entero_o_vacio(df['ano']),
            'mes': self._entero_o_vacio(df['mes']),
            'nombre': df['nombre'] if 'nombre' in df.columns else "",
            'dia': self._entero_o_vacio(df['dia']),
            'detalle': df['detalle'],
            'codpuc': df['codpuc'],
            'tercero': df['tercero'],
            'docrela': df['docrela'],
            'debito': df['debito'].astype(object).where(df['debito'].notna(), ""),
            'credito': df['credito'].astype(object).where(df['credito'].notna(), ""),
        }).reindex(columns=self.COLUMNAS_PLANO, fill_value="")

        ruta_temp_xlsx = os.path.splitext(ruta_destino_final)[0] + ".xlsx"
        df_plano.to_excel(ruta_temp_xlsx, index=False, engine='openpyxl')
        self._convertir_a_excel_95(ruta_temp_xlsx, ruta_destino_final)

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
            partes.append(lin.assign(mov=lin['mov'] + base, archivo=archivo))
            base += int(lin['mov'].max()) + 1
        return pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()