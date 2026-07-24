import os
import re
import pandas as pd
import pdfplumber

DELIMITADOR = "|||"


def es_numero_financiero(texto):
    """Valida si un texto es un monto monetario (positivo, negativo, con comas, puntos o signo $)."""
    if not texto:
        return False
    txt = texto.strip().replace("$", "").replace(" ", "")
    patron = r"^[-+]?[\d\.,]+$"
    return bool(re.match(patron, txt))


def pdf_a_texto_delimitado(pdf_path):
    # Detecta fechas YYYY/MM/DD (2026/06/30), DD/MM/YYYY o D/MM/YY
    patron_fecha = r"^\b(?:\d{4}/\d{2}/\d{2}|\d{1,2}/\d{2}(?:/\d{2,4})?)\b"
    lineas_delimitadas = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            anc_pagina = page.width

            # -------------------------------------------------------------
            # LÍMITES HORIZONTALES PARA LAS COLUMNAS CENTRALES
            # -------------------------------------------------------------
            x_max_desc = anc_pagina * 0.35
            x_max_sucursal = anc_pagina * 0.51
            x_max_ref1 = anc_pagina * 0.64
            x_max_ref2 = anc_pagina * 0.75
            x_max_doc = anc_pagina * 0.85

            # Agrupar palabras por altura vertical Y (renglones)
            filas_y = {}
            for w in words:
                y_key = round(w["top"] / 3.0) * 3.0
                filas_y.setdefault(y_key, []).append(w)

            # Lista de transacciones encontradas en la página
            transacciones_pagina = []
            tx_actual = None

            for y_key in sorted(filas_y.keys()):
                words_linea = sorted(filas_y[y_key], key=lambda x: x["x0"])
                if not words_linea:
                    continue

                primera_word = words_linea[0]
                texto_primera = primera_word["text"].strip()

                # ¿Es el inicio de una nueva transacción? (Tiene fecha en el margen izquierdo)
                es_inicio_tx = bool(
                    re.match(patron_fecha, texto_primera)
                    and primera_word["x0"] < (anc_pagina * 0.20)
                )

                if es_inicio_tx:
                    # Guardar transacción previa si existía
                    if tx_actual:
                        transacciones_pagina.append(tx_actual)

                    fecha_val = texto_primera
                    resto_words = words_linea[1:]

                    # Extraer el valor financiero si está al final del primer renglón
                    valor_val = ""
                    if len(resto_words) > 0 and es_numero_financiero(
                        resto_words[-1]["text"]
                    ):
                        valor_val = resto_words.pop(-1)["text"]

                    # Inicializar estructura de la nueva transacción
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
                    # Es un renglón secundario (continuación de la transacción anterior)
                    if not tx_actual:
                        continue  # Ignorar si es texto de encabezado de página sin transacción previa
                    words_a_procesar = words_linea

                    # Si el renglón secundario trae el monto (por si venía abajo)
                    if not tx_actual["VALOR"] and len(words_a_procesar) > 0:
                        if es_numero_financiero(words_a_procesar[-1]["text"]):
                            tx_actual["VALOR"] = words_a_procesar.pop(-1)[
                                "text"
                            ]

                # Clasificar las palabras en las columnas intermedias según su coordenada X
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
                        tx_actual["DOCUMENTO"].append(txt_w)

            # Guardar la última transacción procesada de la página
            if tx_actual:
                transacciones_pagina.append(tx_actual)

            # Convertir las transacciones acumuladas en texto plano delimitado
            for tx in transacciones_pagina:
                desc_str = " ".join(tx["DESCRIPCIÓN"]).strip()
                suc_str = " ".join(tx["SUCURSAL/CANAL"]).strip()
                ref1_str = " ".join(tx["REFERENCIA 1"]).strip()
                ref2_str = " ".join(tx["REFERENCIA 2"]).strip()
                doc_str = " ".join(tx["DOCUMENTO"]).strip()

                # Ignorar encabezados azules si se capturaron por error
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


def texto_delimitado_a_excel(lineas_texto, output_excel_path):
    columnas = [
        "FECHA",
        "DESCRIPCIÓN",
        "SUCURSAL/CANAL",
        "REFERENCIA 1",
        "REFERENCIA 2",
        "DOCUMENTO",
        "VALOR",
    ]
    registros = []

    for linea in lineas_texto:
        partes = linea.split(DELIMITADOR)
        if len(partes) == len(columnas):
            registros.append({
                "FECHA": partes[0],
                "DESCRIPCIÓN": partes[1],
                "SUCURSAL/CANAL": partes[2],
                "REFERENCIA 1": partes[3],
                "REFERENCIA 2": partes[4],
                "DOCUMENTO": partes[5],
                "VALOR": partes[6],
            })

    df = pd.DataFrame(registros, columns=columnas)
    df.to_excel(output_excel_path, index=False)


def reorganizar_excel(excel_path):
    """Reorganiza las filas del archivo Excel de modo que los valores con '-' queden en la parte superior."""
    if not os.path.exists(excel_path):
        return

    df = pd.read_excel(excel_path, dtype=str)

    if "VALOR" not in df.columns or df.empty:
        return

    is_negative = df["VALOR"].astype(str).str.strip().str.contains("-", na=False)

    df["_orden_temp"] = 0
    df.loc[~is_negative, "_orden_temp"] = 1

    df_ordenado = df.sort_values(by="_orden_temp", kind="stable").drop(
        columns=["_orden_temp"]
    )
    df_ordenado.to_excel(excel_path, index=False)


def ejecutar_proceso_exportacion(pdf_path):
    base_path, _ = os.path.splitext(pdf_path)
    txt_path = f"{base_path}_plano.txt"
    excel_path = f"{base_path}_convertido.xlsx"

    lineas_plana = pdf_a_texto_delimitado(pdf_path)

    if not lineas_plana:
        print("⚠ ATENCIÓN: No se encontraron movimientos en el PDF.")
        return

    encabezado_txt = DELIMITADOR.join([
        "FECHA",
        "DESCRIPCIÓN",
        "SUCURSAL/CANAL",
        "REFERENCIA 1",
        "REFERENCIA 2",
        "DOCUMENTO",
        "VALOR",
    ])
    contenido_txt = [encabezado_txt] + lineas_plana

    # 1. Guardar archivo Plano (.txt)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(contenido_txt))

    # 2. Guardar archivo Excel (.xlsx)
    texto_delimitado_a_excel(lineas_plana, excel_path)

    # 3. Reorganizar filas en Excel
    reorganizar_excel(excel_path)

    print(
        f"✓ Éxito: Se procesaron {len(lineas_plana)} transacciones completas."
    )
    print(f"✓ Archivo Plano exportado en: {txt_path}")
    print(f"✓ Archivo Excel exportado en: {excel_path}")


if __name__ == "__main__":
    archivo_pdf = r"C:\Users\USUARIO\Desktop\ZIP_02138656805_000000900823072_20260707_12230941.pdf"
    ejecutar_proceso_exportacion(archivo_pdf)