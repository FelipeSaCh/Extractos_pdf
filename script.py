import os
import re
import pandas as pd
import pdfplumber

DELIMITADOR = "|||"


def es_numero_financiero(texto):
    """Valida si un texto es un monto monetario (positivo, negativo o decimales)."""
    txt = texto.strip().replace("$", "")
    patron = r"^-?[\d\.,]+$"
    return bool(re.match(patron, txt))


def pdf_a_texto_delimitado(pdf_path):
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

                if re.match(patron_fecha, texto_primera) and primera_word["x0"] < (anc_pagina * 0.20):
                    fecha_val = texto_primera
                    resto_words = words_linea[1:]
                    if not resto_words:
                        continue

                    saldo_val = ""
                    valor_val = ""

                    if len(resto_words) > 0 and es_numero_financiero(resto_words[-1]["text"]):
                        saldo_val = resto_words.pop(-1)["text"]

                    if len(resto_words) > 0 and es_numero_financiero(resto_words[-1]["text"]):
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
                        elif x_centro < x_max_dcto:
                            dcto_words.append(txt_w)
                        else:
                            dcto_words.append(txt_w)

                    descripcion_val = " ".join(descripcion_words).strip()
                    sucursal_val = " ".join(sucursal_words).strip()
                    dcto_val = " ".join(dcto_words).strip()

                    if "DESCRIPC" in descripcion_val.upper() or "RESUMEN" in descripcion_val.upper():
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


def texto_delimitado_a_excel(lineas_texto, output_excel_path):
    columnas = ["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO.", "VALOR", "SALDO"]
    registros = []

    for linea in lineas_texto:
        partes = linea.split(DELIMITADOR)
        if len(partes) == len(columnas):
            registros.append({
                "FECHA": partes[0],
                "DESCRIPCIÓN": partes[1],
                "SUCURSAL": partes[2],
                "DCTO.": partes[3],
                "VALOR": partes[4],
                "SALDO": partes[5],
            })

    df = pd.DataFrame(registros, columns=columnas)
    df.to_excel(output_excel_path, index=False)


def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None):
    """
    Procesa el PDF y guarda el Excel (.xlsx) y el archivo plano (.txt).
    Si no se proporciona 'output_excel_path', se asume el nombre predeterminado.
    """
    if not output_excel_path:
        base_path, _ = os.path.splitext(pdf_path)
        output_excel_path = f"{base_path}_convertido.xlsx"

    base_txt_path, _ = os.path.splitext(output_excel_path)
    txt_path = f"{base_txt_path}_plano.txt"

    lineas_plana = pdf_a_texto_delimitado(pdf_path)

    encabezado_txt = DELIMITADOR.join([
        "FECHA",
        "DESCRIPCIÓN",
        "SUCURSAL",
        "DCTO.",
        "VALOR",
        "SALDO",
    ])
    contenido_txt = [encabezado_txt] + lineas_plana

    # 1. Guardar archivo Plano (.txt)
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(contenido_txt))

    # 2. Guardar archivo Excel (.xlsx)
    texto_delimitado_a_excel(lineas_plana, output_excel_path)

    return output_excel_path


def reorganizar_excel(excel_path):
    """
    Reorganiza las filas del archivo Excel de modo que los valores negativos o 
    que comiencen con '-' en la columna 'VALOR' queden en la parte superior.
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"No se encontró el archivo: {excel_path}")

    df = pd.read_excel(excel_path, dtype=str)

    if "VALOR" not in df.columns:
        raise KeyError("La columna 'VALOR' no existe en el archivo Excel.")

    is_negative = df["VALOR"].astype(str).str.strip().str.contains("-", na=False)

    df["_orden_temp"] = 0
    df.loc[~is_negative, "_orden_temp"] = 1

    df_ordenado = df.sort_values(by="_orden_temp", kind="stable").drop(columns=["_orden_temp"])
    df_ordenado.to_excel(excel_path, index=False)
    return excel_path


if __name__ == "__main__":
    archivo_pdf = r"C:\Users\USUARIO\Downloads\51962077_Cuentasdeahorro_8538_unlocked.pdf"
    ejecutar_proceso_exportacion(archivo_pdf)