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

    CONCEPTOS_BANCARIOS = [
        ("GMF 4X1000", r"4\s*X\s*1\.?000"),
        ("IVA", r"\bIVA\b"),
        ("CUOTA PLAN CANAL NEGOCIOS", r"CUOTA\s+PLAN\s+CANAL\s+NEGOCIOS"),
        ("TARJETA DE CREDITO", r"TARJETA\s+DE\s+CR[EÉ]DITO"),
        ("COMISION", r"COMISI[OÓ]N"),
        ("SERVICIO", r"SERVICIO"),
        ("TARJ DEB", r"TARJ\s+DEB"),
        ("INTERESES AHORROS", r"INTERESES\s+AHORROS"),
    ]
    CODPUC_BANCARIOS = {
        "GMF 4X1000": "530595001",
        "IVA": "530505002",
        "CUOTA PLAN CANAL NEGOCIOS": "530505001",
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
    PATRON_RETIROS = r"\bRETIROS?\b"

    CLASE_EGRESO = "EGRESO"
    CLASE_NOMINA = "NOMINA"
    CLASE_RECIBO = "RECIBO"
    CLASE_BANCARIO = "BANCARIO"

    COMPROBANTES_PREDETERMINADOS = {
        CLASE_EGRESO: "EG1",
        CLASE_NOMINA: "EG2",
        CLASE_RECIBO: "RC1",
        CLASE_BANCARIO: "CC1",
    }
    ORDEN_PREFIJOS = {"CC": 0, "EG": 1, "RC": 2}

    LADO_BANCO = "BANCO"
    LADO_CONTRA = "CONTRA"

    PATRON_NOMINA = r"PAGO A NOMIN"
    DETALLE_NOMINA = "PAGO A NOMIN"

    MESES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO",
             "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

    def __init__(self, ruta_excel_origen):
        self.ruta_origen = ruta_excel_origen

    def _limpiar_y_convertir_valores(self, df):
        if 'VALOR' not in df.columns:
            return pd.Series(0.0, index=df.index)
        if df['VALOR'].dtype == object:
            valor_limpio = df['VALOR'].astype(str).str.replace(r'[^\d\.-]', '', regex=True)
            return pd.to_numeric(valor_limpio, errors='coerce').fillna(0.0)
        return df['VALOR'].fillna(0.0)

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

        fechas_dt = pd.to_datetime(df['FECHA'], errors='coerce', dayfirst=True, format="mixed")
        texto = df['FECHA'].astype(str)

        dia = pd.to_numeric(fechas_dt.dt.day.fillna(texto.str.extract(r'^(\d{1,2})')[0]), errors='coerce')
        mes = pd.to_numeric(fechas_dt.dt.month.fillna(texto.str.extract(r'[/.-](\d{1,2})')[0]), errors='coerce')
        ano = pd.to_numeric(
            fechas_dt.dt.year.replace(1900, np.nan).fillna(texto.str.extract(r'[/.-]\d{1,2}[/.-](\d{4})')[0]),
            errors='coerce'
        )
        return ano.where(ano >= 2000), mes, dia

    def _construir_movimientos(self, df):
        n = len(df)
        valores = self._limpiar_y_convertir_valores(df)
        ano, mes, dia = self._extraer_fechas(df)

        desc = df.get("DESCRIPCIÓN", df.get("DESCRIPCION", pd.Series([""] * n, index=df.index)))
        desc = desc.fillna("").astype(str).str.strip()

        col_ref = next((c for c in ["REFERENCIA 1", "REFERENCIA", "DCTO."] if c in df.columns), None)
        if col_ref:
            ref = df[col_ref].fillna("").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            detalle = (desc + " " + ref).str.strip()
            tercero = ref.apply(self._es_tercero_valido)
        else:
            detalle = desc
            tercero = pd.Series("", index=df.index)

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

        agrupado = (
            mov[mask]
            .groupby(['ano', 'mes', 'dia'], as_index=False, dropna=False)['valor'].sum()
            .assign(
                clase=self.CLASE_NOMINA,
                desc="",
                detalle=self.DETALLE_NOMINA,
                tercero="",
                codpuc_contra="",
            )
        )
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
        return lin.assign(codpuc_manual="", tercero_manual="", docrela="")

    def calcular(self, lin, nit_empresa="", primer_consecutivo=1):
        df = lin.copy()
        for col in ['codpuc_contra', 'tercero_ref', 'codpuc_manual', 'tercero_manual', 'docrela']:
            df[col] = df[col].fillna("").astype(str)
        df['tipo'] = df['tipo'].fillna("").astype(str).str.strip().str.upper()

        prefijo = df['tipo'].str[:2]
        entrada = df['valor'] > 0
        es_banco = df['lado'] == self.LADO_BANCO
        en_debito = es_banco == entrada

        contra = np.select(
            [
                prefijo == "CC",
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
            df.assign(_rango=prefijo.map(self.ORDEN_PREFIJOS).fillna(len(self.ORDEN_PREFIJOS)), _credito=~en_debito)
            .sort_values(['_rango', 'tipo', 'detalle', 'ano', 'mes', 'dia', 'mov', '_credito'],
                         kind='stable', na_position='last')
            .drop(columns=['_rango', '_credito'])
        )
        inicio = self._a_entero(primer_consecutivo)
        df['numero'] = [f"{n:010d}" for n in pd.factorize(df['mov'])[0] + inicio]
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
            'dia': self._entero_o_vacio(df['dia']),
            'detalle': df['detalle'],
            'codpuc': df['codpuc'],
            'tercero': df['tercero'],
            'docrela': df['docrela'],
            'debito': df['debito'].astype(object).where(df['debito'].notna(), ""),
            'credito': df['credito'].astype(object).where(df['credito'].notna(), ""),
        }).reindex(columns=self.COLUMNAS_PLANO, fill_value="")

        ruta_temp_xlsx = ruta_destino_final.replace(".xls", ".xlsx")
        df_plano.to_excel(ruta_temp_xlsx, index=False, engine='openpyxl')
        self._convertir_a_excel_95(ruta_temp_xlsx, ruta_destino_final)

        return ruta_destino_final