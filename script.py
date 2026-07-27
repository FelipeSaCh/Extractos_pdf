import os
import re
import pandas as pd
import pdfplumber

DELIMITADOR = "|||"

# ==============================================================================
# FUNCIONES AUXILIARES COMUNES
# ==============================================================================

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
            x_max_sucursal = anc_pagina * 0.51
            x_max_ref1 = anc_pagina * 0.64
            x_max_ref2 = anc_pagina * 0.75
            x_max_doc = anc_pagina * 0.85

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

                    fecha_val = texto_primera
                    resto_words = words_linea[1:]

                    valor_val = ""
                    if len(resto_words) > 0 and es_numero_financiero(
                        resto_words[-1]["text"]
                    ):
                        valor_val = resto_words.pop(-1)["text"]

                    tx_actual = {
                        "FECHA": fecha_val,
                        "DESCRIPCIÓN": [],
                        "SUCURSAL/CANAL": [],
                        "REFERENCIA 1": [],
                        "REFERENCIA 2": [],
                        "DOCUMENTO": [],
                        "VALOR": valor_val,
                    }
                    words_a_procesar = resto_words

                else:
                    if not tx_actual:
                        continue
                    words_a_procesar = words_linea

                    if not tx_actual["VALOR"] and len(words_a_procesar) > 0:
                        if es_numero_financiero(words_a_procesar[-1]["text"]):
                            tx_actual["VALOR"] = words_a_procesar.pop(-1)[
                                "text"
                            ]

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
                    else:
                        tx_actual["DOCUMENTO"].append(txt_w)

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

    # 4. Agrupar por Conceptos
    df_conceptos = (
        df_ordenado.groupby(col_desc, as_index=False)["VALOR"]
        .sum()
        .rename(columns={"VALOR": "VALOR_TOTAL"})
    )

    # 5. Exportación multi-hoja
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_ordenado.to_excel(writer, sheet_name="Datos", index=False)
        df_conciliacion.to_excel(writer, sheet_name="Resumen", index=False)
        df_conceptos.to_excel(writer, sheet_name="Conceptos", index=False)

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
        print(f"⚠ ATENCIÓN: No se encontraron registros para {tipo}.")
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

    print(f"✓ Éxito [{tipo}]: Procesadas {len(lineas_plana)} transacciones.")
    print(f"✓ Archivo Plano: {txt_path}")
    print(f"✓ Archivo Excel: {output_excel_path}")

    return output_excel_path