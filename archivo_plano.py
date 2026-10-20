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

    TIPO_EGRESO = "EG1"
    TIPO_NOMINA = "EG2"
    TIPO_RECIBO = "RC1"
    TIPO_BANCARIOS = "CC1"

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

    def _construir_movimientos(self, df, nit_empresa):
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

        if nit_empresa:
            tercero = tercero.replace("", nit_empresa)

        mov = pd.DataFrame({
            'tipo': "",
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

    def _consolidar_bancarios(self, mov, nit_empresa):
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
            'tipo': self.TIPO_BANCARIOS,
            'ano': fin_mes.dt.year,
            'mes': fin_mes.dt.month,
            'dia': fin_mes.dt.day,
            'desc': "",
            'detalle': grp['concepto'] + " " + fin_mes.dt.month.map(lambda m: self.MESES[m - 1]) + " " + fin_mes.dt.year.astype(str),
            'tercero': nit_empresa,
            'valor': grp['valor'],
            'codpuc_contra': grp['concepto'].map(self.CODPUC_BANCARIOS).fillna(""),
        })

        return pd.concat([mov[~mask], consolidado], ignore_index=True)

    def _agrupar_nomina(self, mov, nit_empresa):
        mask = (mov['tipo'] == "") & mov['detalle'].str.contains(self.PATRON_NOMINA, case=False, regex=True, na=False)
        if not mask.any():
            return mov

        agrupado = (
            mov[mask]
            .groupby(['ano', 'mes', 'dia'], as_index=False, dropna=False)['valor'].sum()
            .assign(
                tipo=self.TIPO_NOMINA,
                desc="",
                detalle=self.DETALLE_NOMINA,
                tercero=nit_empresa,
                codpuc_contra="",
            )
        )
        agrupado = agrupado[agrupado['valor'] != 0]

        return pd.concat([mov[~mask], agrupado], ignore_index=True)

    def _asignar_tipos_generales(self, mov):
        pendiente = mov['tipo'] == ""
        mov.loc[pendiente, 'tipo'] = np.where(
            mov.loc[pendiente, 'valor'] < 0, self.TIPO_EGRESO, self.TIPO_RECIBO
        )
        return mov

    def _asignar_codpuc(self, mov, en_debito, es_natural):
        es_bancario = mov['tipo'] == self.TIPO_BANCARIOS
        clase = mov['tipo'].str[:2]
        return np.select(
            [
                es_bancario & es_natural,
                es_bancario,
                (clase == "EG") & en_debito,
                (clase == "RC") & ~en_debito,
            ],
            [mov['codpuc_contra'], self.CODPUC_BANCO, self.CODPUC_PROVEEDORES, self.CODPUC_CLIENTES],
            default=self.CODPUC_BANCO,
        )

    def _construir_lineas(self, mov, en_debito, es_natural):
        importe = mov['valor'].abs()
        return pd.DataFrame({
            'tipo': mov['tipo'],
            'numero': mov['numero'],
            'ano': self._entero_o_vacio(mov['ano']),
            'mes': self._entero_o_vacio(mov['mes']),
            'dia': self._entero_o_vacio(mov['dia']),
            'detalle': mov['detalle'],
            'codpuc': self._asignar_codpuc(mov, en_debito, es_natural),
            'tercero': mov['tercero'],
            'debito': importe.where(en_debito, ""),
            'credito': importe.where(~en_debito, ""),
        })

    def _generar_lineas(self, mov, numero_inicial):
        mov = mov.sort_values(['ano', 'mes', 'dia'], kind='stable', na_position='last').reset_index(drop=True)
        mov['numero'] = [f"{n:010d}" for n in range(numero_inicial, numero_inicial + len(mov))]

        salida = mov['valor'] < 0
        es_bancario = mov['tipo'] == self.TIPO_BANCARIOS
        en_debito = pd.Series(np.where(es_bancario, salida, ~salida), index=mov.index)
        doble = mov['tipo'] != self.TIPO_NOMINA

        natural = self._construir_lineas(mov, en_debito, True)
        contrapartida = self._construir_lineas(mov[doble], ~en_debito[doble], False)

        lineas = pd.concat([natural, contrapartida])
        lineas = (
            lineas.assign(_comprobante=lineas.index, _credito=lineas['debito'].eq(""))
            .sort_values(['_comprobante', '_credito'], kind='stable')
            .drop(columns=['_comprobante', '_credito'])
            .reset_index(drop=True)
        )
        return lineas.reindex(columns=self.COLUMNAS_PLANO, fill_value="")

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

    def generar(self, numero_inicial="1", nit_empresa="", ruta_destino_final=None):
        if not ruta_destino_final:
            base, _ = os.path.splitext(self.ruta_origen)
            ruta_destino_final = f"{base}_PLANO.xls"

        try:
            df = pd.read_excel(self.ruta_origen, sheet_name="Datos")
        except ValueError:
            raise ValueError("El archivo seleccionado no contiene la hoja 'Datos'.")

        if df.empty:
            return ruta_destino_final

        try:
            inicio = int(str(numero_inicial).strip() or "1")
        except ValueError:
            inicio = 1

        nit_empresa = str(nit_empresa or "").strip()

        mov = self._construir_movimientos(df, nit_empresa)
        mov = self._consolidar_bancarios(mov, nit_empresa)
        mov = self._agrupar_nomina(mov, nit_empresa)
        mov = self._asignar_tipos_generales(mov)
        df_plano = self._generar_lineas(mov, inicio)

        ruta_temp_xlsx = ruta_destino_final.replace(".xls", ".xlsx")
        df_plano.to_excel(ruta_temp_xlsx, index=False, engine='openpyxl')
        self._convertir_a_excel_95(ruta_temp_xlsx, ruta_destino_final)

        return ruta_destino_final