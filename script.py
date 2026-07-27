import os
import re
import pandas as pd
import pdfplumber
import sys
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DELIMITADOR = "|||"

# ==============================================================================
# FUNCIONES AUXILIARES COMUNES
# ==============================================================================
def obtener_ruta_asset(ruta_relativa):
    """ Obtiene la ruta absoluta para assets, funciona en desarrollo y en el .exe de PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, ruta_relativa)
    return os.path.join(os.path.abspath("."), ruta_relativa)

def es_numero_financiero(texto):
    """Valida si un texto es un monto monetario (positivo, negativo o decimales)."""
    if not texto:
        return False
    txt = texto.strip().replace("$", "").replace(" ", "")
    patron = r"^-?[\d\.,]+$"
    return bool(re.match(patron, txt))


def texto_delimitado_a_excel(lineas_texto, columnas, output_excel_path):
    """Convierte las líneas delimitadas a un DataFrame y lo exporta a Excel."""
    registros = []
    for linea in lineas_texto:
        partes = linea.split(DELIMITADOR)
        if len(partes) == len(columnas):
            registros.append(dict(zip(columnas, partes)))

    df = pd.DataFrame(registros, columns=columnas)
    df.to_excel(output_excel_path, index=False)


# ==============================================================================
# PARSER 1: MOVIMIENTOS
# ==============================================================================

def extraer_lineas_movimientos(pdf_path):
    """Extrae las transacciones para archivos de tipo 'Movimientos'."""
    patron_fecha = r"^\b(?:\d{4}/\d{2}/\d{2}|\d{1,2}/\d{2}(?:/\d{2,4})?)\b"
    lineas_delimitadas = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            anc_pagina = page.width
            x_max_desc = anc_pagina * 0.35
            x_max_sucursal = anc_pagina * 0.50
            x_max_ref1 = anc_pagina * 0.63
            x_max_ref2 = anc_pagina * 0.73
            x_max_doc = anc_pagina * 0.85  # Todo lo que esté a la derecha de 0.85 es la columna VALOR

            filas_y = {}
            for w in words:
                y_key = round(w["top"] / 3.0) * 3.0
                filas_y.setdefault(y_key, []).append(w)

            transacciones_pagina = []
            tx_actual = None

            for y_key in sorted(filas_y.keys()):
                words_linea = sorted(filas_y[y_key], key=lambda x: x["x0"])
                if not words_linea:
                    continue

                primera_word = words_linea[0]
                texto_primera = primera_word["text"].strip()

                es_inicio_tx = bool(
                    re.match(patron_fecha, texto_primera)
                    and primera_word["x0"] < (anc_pagina * 0.20)
                )

                if es_inicio_tx:
                    if tx_actual:
                        transacciones_pagina.append(tx_actual)

                    tx_actual = {
                        "FECHA": texto_primera,
                        "DESCRIPCIÓN": [],
                        "SUCURSAL/CANAL": [],
                        "REFERENCIA 1": [],
                        "REFERENCIA 2": [],
                        "DOCUMENTO": [],
                        "VALOR": "",
                    }
                    words_a_procesar = words_linea[1:]
                else:
                    if not tx_actual:
                        continue
                    words_a_procesar = words_linea

                # Asignación estricta por coordenadas X
                for w in words_a_procesar:
                    x_centro = (w["x0"] + w["x1"]) / 2.0
                    txt_w = w["text"]

                    if x_centro < x_max_desc:
                        tx_actual["DESCRIPCIÓN"].append(txt_w)
                    elif x_centro < x_max_sucursal:
                        tx_actual["SUCURSAL/CANAL"].append(txt_w)
                    elif x_centro < x_max_ref1:
                        tx_actual["REFERENCIA 1"].append(txt_w)
                    elif x_centro < x_max_ref2:
                        tx_actual["REFERENCIA 2"].append(txt_w)
                    elif x_centro < x_max_doc:
                        tx_actual["DOCUMENTO"].append(txt_w)
                    else:
                        # Columna final: VALOR
                        tx_actual["VALOR"] = txt_w

            if tx_actual:
                transacciones_pagina.append(tx_actual)

            for tx in transacciones_pagina:
                desc_str = " ".join(tx["DESCRIPCIÓN"]).strip()
                suc_str = " ".join(tx["SUCURSAL/CANAL"]).strip()
                ref1_str = " ".join(tx["REFERENCIA 1"]).strip()
                ref2_str = " ".join(tx["REFERENCIA 2"]).strip()
                doc_str = " ".join(tx["DOCUMENTO"]).strip()

                if (
                    "DESCRIPC" in desc_str.upper()
                    or "REFERENCIA" in ref1_str.upper()
                ):
                    continue

                registro = [
                    tx["FECHA"],
                    desc_str,
                    suc_str,
                    ref1_str,
                    ref2_str,
                    doc_str,
                    tx["VALOR"],
                ]
                lineas_delimitadas.append(DELIMITADOR.join(registro))

    return lineas_delimitadas

# ==============================================================================
# PARSER 2: EXTRACTOS
# ==============================================================================

def extraer_lineas_extractos(pdf_path):
    """Extrae las transacciones para archivos de tipo 'Extracto'."""
    patron_fecha = r"^\b\d{1,2}/\d{2}(?:/\d{2,4})?\b"
    lineas_delimitadas = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            anc_pagina = page.width
            x_max_desc = anc_pagina * 0.40
            x_max_sucursal = anc_pagina * 0.65
            x_max_dcto = anc_pagina * 0.78

            filas_y = {}
            for w in words:
                y_key = round(w["top"] / 2.5) * 2.5
                filas_y.setdefault(y_key, []).append(w)

            for y_key in sorted(filas_y.keys()):
                words_linea = sorted(filas_y[y_key], key=lambda x: x["x0"])
                if not words_linea:
                    continue

                primera_word = words_linea[0]
                texto_primera = primera_word["text"].strip()

                if re.match(patron_fecha, texto_primera) and primera_word[
                    "x0"
                ] < (anc_pagina * 0.20):
                    fecha_val = texto_primera
                    resto_words = words_linea[1:]
                    if not resto_words:
                        continue

                    saldo_val = ""
                    valor_val = ""

                    if len(resto_words) > 0 and es_numero_financiero(
                        resto_words[-1]["text"]
                    ):
                        saldo_val = resto_words.pop(-1)["text"]

                    if len(resto_words) > 0 and es_numero_financiero(
                        resto_words[-1]["text"]
                    ):
                        valor_val = resto_words.pop(-1)["text"]

                    descripcion_words = []
                    sucursal_words = []
                    dcto_words = []

                    for w in resto_words:
                        x_centro = (w["x0"] + w["x1"]) / 2.0
                        txt_w = w["text"]

                        if x_centro < x_max_desc:
                            descripcion_words.append(txt_w)
                        elif x_centro < x_max_sucursal:
                            sucursal_words.append(txt_w)
                        else:
                            dcto_words.append(txt_w)

                    descripcion_val = " ".join(descripcion_words).strip()
                    sucursal_val = " ".join(sucursal_words).strip()
                    dcto_val = " ".join(dcto_words).strip()

                    if (
                        "DESCRIPC" in descripcion_val.upper()
                        or "RESUMEN" in descripcion_val.upper()
                    ):
                        continue

                    registro = [
                        fecha_val,
                        descripcion_val,
                        sucursal_val,
                        dcto_val,
                        valor_val,
                        saldo_val,
                    ]
                    lineas_delimitadas.append(DELIMITADOR.join(registro))

    return lineas_delimitadas


# ==============================================================================
# REORGANIZACIÓN Y GENERACIÓN DE REPORTES EN EXCEL
# ==============================================================================

def reorganizar_excel(excel_path):
    """Limpia tipos de datos, ordena negativos arriba y genera las hojas

    'Datos', 'Resumen' y 'Conceptos'.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")

    df = pd.read_excel(excel_path, dtype=str)

    if "VALOR" not in df.columns or df.empty:
        return excel_path

    col_desc = "DESCRIPCIÓN" if "DESCRIPCIÓN" in df.columns else "DESCRIPCION"

    # 1. Identificar negativos y ordenar
    is_negative = (
        df["VALOR"].astype(str).str.strip().str.contains("-", na=False)
    )
    df["_orden_temp"] = 0
    df.loc[~is_negative, "_orden_temp"] = 1

    df_ordenado = df.sort_values(by="_orden_temp", kind="stable").drop(
        columns=["_orden_temp"]
    )

    # 2. Formatear números
    df_ordenado["VALOR"] = pd.to_numeric(
        df_ordenado["VALOR"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.strip(),
        errors="coerce",
    )

    if "SALDO" in df_ordenado.columns:
        df_ordenado["SALDO"] = pd.to_numeric(
            df_ordenado["SALDO"]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.strip(),
            errors="coerce",
        )

    # 3. Conciliación (Resumen)
    df_conciliacion = pd.DataFrame(
        {
            "CARGOS": [df_ordenado.loc[df_ordenado["VALOR"] < 0, "VALOR"].sum()],
            "ABONOS": [df_ordenado.loc[df_ordenado["VALOR"] > 0, "VALOR"].sum()],
        }
    )

   # ==============================================================================
    # 4. Agrupar por Conceptos (Separando Cargos y Abonos)
    # ==============================================================================

    # Crear columnas auxiliares para separar montos según su signo
    df_ordenado["CARGOS_TEMP"] = df_ordenado["VALOR"].apply(
        lambda x: x if x < 0 else 0
    )
    df_ordenado["ABONOS_TEMP"] = df_ordenado["VALOR"].apply(
        lambda x: x if x > 0 else 0
    )

    # Agrupar sumando por separado cada tipo de movimiento
    df_conceptos = (
        df_ordenado.groupby(col_desc, as_index=False)
        .agg(
            CARGOS=("CARGOS_TEMP", "sum"),
            ABONOS=("ABONOS_TEMP", "sum"),
            NETO=("VALOR", "sum"),
        )
        .sort_values(by="CARGOS", ascending=True)  # Ordena mostrando mayor gasto arriba
    )

    # Limpiar columnas auxiliares creadas en el DataFrame principal
    df_ordenado.drop(columns=["CARGOS_TEMP", "ABONOS_TEMP"], inplace=True)

    # ==============================================================================
    # 5. Exportación multi-hoja con Formato de Moneda y Fila de TOTALES
    # ==============================================================================
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_ordenado.to_excel(writer, sheet_name="Datos", index=False)
        df_conciliacion.to_excel(writer, sheet_name="Resumen", index=False)
        df_conceptos.to_excel(writer, sheet_name="Conceptos", index=False)

        workbook = writer.book

        # Formato de moneda contable ($)
        FORMATO_MONEDA = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'

        # Estilos contables para la fila de TOTAL
        font_bold = Font(bold=True)
        fill_total = PatternFill(
            start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
        )
        border_top_thin = Side(border_style="thin", color="000000")
        border_bottom_double = Side(border_style="double", color="000000")
        border_total = Border(top=border_top_thin, bottom=border_bottom_double)

        # --------------------------------------------------------------------------
        # A. Aplicar formato numérico a todas las hojas
        # --------------------------------------------------------------------------
        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]

            for col in worksheet.iter_cols(1, worksheet.max_column):
                header_val = str(col[0].value).upper() if col[0].value else ""

                if header_val in [
                    "VALOR",
                    "SALDO",
                    "CARGOS",
                    "ABONOS",
                    "NETO",
                    "VALOR_TOTAL",
                ]:
                    for cell in col[1:]:
                        if cell.value is not None and isinstance(
                            cell.value, (int, float)
                        ):
                            cell.number_format = FORMATO_MONEDA

        # --------------------------------------------------------------------------
        # B. Agregar fila de TOTAL en la hoja 'Conceptos'
        # --------------------------------------------------------------------------
        if "Conceptos" in workbook.sheetnames:
            ws_conceptos = workbook["Conceptos"]
            max_row = ws_conceptos.max_row
            total_row = max_row + 1

            # Mapear encabezados para ubicar las columnas dinámicamente
            headers = [
                str(cell.value).upper() if cell.value else ""
                for cell in ws_conceptos[1]
            ]

            # Obtener los índices de las columnas
            col_desc_idx = (
                headers.index("DESCRIPCION") + 1 if "DESCRIPCION" in headers else 1
            )
            col_cargos_idx = (
                headers.index("CARGOS") + 1 if "CARGOS" in headers else None
            )
            col_abonos_idx = (
                headers.index("ABONOS") + 1 if "ABONOS" in headers else None
            )
            col_neto_idx = (
                headers.index("NETO") + 1 if "NETO" in headers else None
            )

            # 1. Escribir la etiqueta "TOTAL"
            cell_label = ws_conceptos.cell(
                row=total_row, column=col_desc_idx, value="TOTAL"
            )
            cell_label.font = font_bold
            cell_label.alignment = Alignment(horizontal="left")

            # 2. Insertar fórmulas =SUM(...) para CARGOS, ABONOS y NETO
            for col_idx in [col_cargos_idx, col_abonos_idx, col_neto_idx]:
                if col_idx:
                    col_letter = get_column_letter(col_idx)
                    # Fórmula de Excel para sumar desde el primer dato hasta el último
                    formula = f"=SUM({col_letter}2:{col_letter}{max_row})"

                    cell_total = ws_conceptos.cell(
                        row=total_row, column=col_idx, value=formula
                    )
                    cell_total.font = font_bold
                    cell_total.number_format = FORMATO_MONEDA
                    cell_total.fill = fill_total
                    cell_total.border = border_total

    return excel_path


# ==============================================================================
# ORQUESTADOR PRINCIPAL (MÉTODO ÚNICO DE ENTRADA)
# ==============================================================================

def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None):
    """Detecta automáticamente si el PDF es de 'Movimientos' o 'Extracto' según el nombre del archivo y ejecuta la extracción correspondiente."""
    nombre_archivo = os.path.basename(pdf_path).upper()

    if "MOVIMIENTO" in nombre_archivo:
        tipo = "MOVIMIENTOS"
        columnas = [
            "FECHA",
            "DESCRIPCIÓN",
            "SUCURSAL/CANAL",
            "REFERENCIA 1",
            "REFERENCIA 2",
            "DOCUMENTO",
            "VALOR",
        ]
        lineas_plana = extraer_lineas_movimientos(pdf_path)

    elif "EXTRACTO" in nombre_archivo:
        tipo = "EXTRACTOS"
        columnas = [
            "FECHA",
            "DESCRIPCION",
            "SUCURSAL",
            "DCTO.",
            "VALOR",
            "SALDO",
        ]
        lineas_plana = extraer_lineas_extractos(pdf_path)

    else:
        raise ValueError(
            "El nombre del archivo no contiene la palabra 'Movimientos' ni 'Extracto'."
        )

    if not lineas_plana:
        
        return None

    # Definir rutas de salida
    if not output_excel_path:
        base_path, _ = os.path.splitext(pdf_path)
        output_excel_path = f"{base_path}_convertido.xlsx"

    base_txt_path, _ = os.path.splitext(output_excel_path)
    txt_path = f"{base_txt_path}_plano.txt"

    # 1. Guardar archivo Plano (.txt)
    encabezado_txt = DELIMITADOR.join(columnas)
    contenido_txt = [encabezado_txt] + lineas_plana

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(contenido_txt))

    # 2. Convertir a Excel
    texto_delimitado_a_excel(lineas_plana, columnas, output_excel_path)

    # 3. Reorganizar y crear pestañas suplementarias
    reorganizar_excel(output_excel_path)



    return output_excel_path