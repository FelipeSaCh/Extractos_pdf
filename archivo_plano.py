import pandas as pd
import numpy as np
import re
import os
import win32com.client as win32
import pythoncom

class GeneradorArchivoPlano:
    COLUMNAS_PLANO = [
        "libro", "tipo", "numero", "ano", "mes", "dia", "detalle", "anulado",
        "ccosto", "sccosto", "docrela", "cuota", "artrela", "artcant", "base",
        "base2", "fechaven", "area", "numexterno", "anexo", "bodega", "codpuc",
        "tercero", "debito", "credito", "nombre"
    ]

    def __init__(self, ruta_excel_origen):
        self.ruta_origen = ruta_excel_origen

    def _limpiar_y_convertir_valores(self, df):
        if 'VALOR' not in df.columns:
            return pd.Series([0.0] * len(df))
            
        if df['VALOR'].dtype == object:
            valor_limpio = df['VALOR'].astype(str).str.replace(r'[^\d\.-]', '', regex=True)
            return pd.to_numeric(valor_limpio, errors='coerce').fillna(0.0)
        return df['VALOR'].fillna(0.0)

    def _es_tercero_valido(self, texto):
        if pd.isna(texto) or not texto:
            return ""
            
        texto_limpio = str(texto).strip()
        patron_identificacion = r"^\d{6,15}(?:-\d)?$"
        
        if re.match(patron_identificacion, texto_limpio):
            return texto_limpio
        return ""

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
            base, ext = os.path.splitext(self.ruta_origen)
            ruta_destino_final = f"{base}_PLANO.xls"

        try:
            df = pd.read_excel(self.ruta_origen, sheet_name="Datos")
        except ValueError:
            raise ValueError("El archivo seleccionado no contiene la hoja 'Datos'.")

        df_plano = pd.DataFrame(columns=self.COLUMNAS_PLANO)

        filas_totales = len(df)
        if filas_totales == 0:
            return ruta_destino_final

        # 1. Procesar Valores y Tipos Base
        valores_numericos = self._limpiar_y_convertir_valores(df)
        es_negativo = valores_numericos < 0
        
        df_plano['tipo'] = np.where(es_negativo, 'RC1', 'EG1')
        df_plano['credito'] = np.where(es_negativo, valores_numericos.abs(), "")
        df_plano['debito'] = np.where(~es_negativo, valores_numericos.abs(), "")

        # 2. Procesar Fechas
        if 'FECHA' in df.columns:
            fechas_dt = pd.to_datetime(df['FECHA'], errors='coerce', dayfirst=True, format="mixed")
            serie_fechas_str = df['FECHA'].astype(str)
            
            df_plano['dia'] = fechas_dt.dt.day.fillna(serie_fechas_str.str.extract(r'^(\d{1,2})')[0]).fillna("")
            df_plano['mes'] = fechas_dt.dt.month.fillna(serie_fechas_str.str.extract(r'[/.-](\d{1,2})')[0]).fillna("")
            
            anos_extraidos = fechas_dt.dt.year.replace(1900, np.nan)
            df_plano['ano'] = anos_extraidos.fillna("").apply(lambda x: str(int(x)) if isinstance(x, float) and not pd.isna(x) else x)
        
        desc = df.get("DESCRIPCIÓN", df.get("DESCRIPCION", pd.Series([""] * filas_totales))).astype(str).fillna("")
        col_ref = next((col for col in ["REFERENCIA 1", "REFERENCIA", "DCTO."] if col in df.columns), None)
        
        if col_ref:
            ref = df[col_ref].astype(str).fillna("").str.replace("nan", "")
            df_plano['detalle'] = (desc + " " + ref).str.strip()
            df_plano['tercero'] = ref.apply(self._es_tercero_valido)
        else:
            df_plano['detalle'] = desc.str.strip()
            df_plano['tercero'] = "" # Aseguramos inicializarlo en vacío

        # --- NUEVA LÓGICA: RELLENAR TERCEROS VACÍOS CON EL NIT ---
        if nit_empresa:
            nit_str = str(nit_empresa).strip()
            # Reemplaza cualquier cadena vacía o valores nulos con el NIT ingresado
            df_plano['tercero'] = df_plano['tercero'].replace("", nit_str).fillna(nit_str)
        # ---------------------------------------------------------

        df_plano.fillna("", inplace=True)

        # --- 4. REGLAS ESPECIALES PARA EMPRESA 900823072 ---
        if "900823072" in str(self.ruta_origen):
            # Identificar las filas que contienen "PAGO A NOMIN"
            mask_nomin = df_plano['detalle'].str.contains("PAGO A NOMIN", na=False, case=False)

            # Regla de Tipos: Los demás positivos pasan de EG1 a EG2
            df_plano.loc[(df_plano['tipo'] == 'EG1') & (~mask_nomin), 'tipo'] = 'EG2'
            # Forzamos los pagos de nómina a EG1 (independientemente de su signo original)
            df_plano.loc[mask_nomin, 'tipo'] = 'EG1'

            if mask_nomin.any():
                df_nomin = df_plano[mask_nomin].copy()

                # Convertir a flotantes para poder sumarlos matemáticamente
                df_nomin['credito'] = pd.to_numeric(df_nomin['credito'], errors='coerce').fillna(0)
                df_nomin['debito'] = pd.to_numeric(df_nomin['debito'], errors='coerce').fillna(0)

                # Agrupar por año, mes y día, sumando los montos
                agrupado = df_nomin.groupby(['ano', 'mes', 'dia'], as_index=False).agg({
                    'libro': 'first',
                    'tipo': 'first',
                    'numero': 'first',
                    'detalle': lambda x: "PAGO A NOMIN", # Limpiar descripciones residuales
                    'anulado': 'first',
                    'ccosto': 'first',
                    'sccosto': 'first',
                    'docrela': 'first',
                    'cuota': 'first',
                    'artrela': 'first',
                    'artcant': 'first',
                    'base': 'first',
                    'base2': 'first',
                    'fechaven': 'first',
                    'area': 'first',
                    'numexterno': 'first',
                    'anexo': 'first',
                    'bodega': 'first',
                    'codpuc': 'first',
                    'tercero': 'first',
                    'debito': 'sum',
                    'credito': 'sum',
                    'nombre': 'first'
                })

                # Volver a colocar como vacíos los ceros que se hayan sumado sin sentido
                agrupado['credito'] = agrupado['credito'].replace(0.0, "")
                agrupado['debito'] = agrupado['debito'].replace(0.0, "")

                # Remover las filas originales y anexar las agrupadas al final
                df_plano = df_plano[~mask_nomin]
                df_plano = pd.concat([df_plano, agrupado], ignore_index=True)
        # ---------------------------------------------------

        # --- 5. ASIGNACIÓN DEL CONSECUTIVO DE 10 DÍGITOS ---
        # (Se calcula al final por si el número total de filas cambió por agrupación)
        try:
            inicio_int = int(str(numero_inicial).strip() or "1")
        except ValueError:
            inicio_int = 1

        filas_finales = len(df_plano)
        secuencia_numeros = range(inicio_int, inicio_int + filas_finales)
        df_plano['numero'] = [f"{n:010d}" for n in secuencia_numeros]
        # ---------------------------------------------------

        df_plano.fillna("", inplace=True)
        
        ruta_temp_xlsx = ruta_destino_final.replace(".xls", ".xlsx")
        df_plano.to_excel(ruta_temp_xlsx, index=False, engine='openpyxl')
        
        self._convertir_a_excel_95(ruta_temp_xlsx, ruta_destino_final)
        
        return ruta_destino_final