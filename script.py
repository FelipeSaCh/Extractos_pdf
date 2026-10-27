import os
import re

import pdfplumber

from comun_extractos import (
    DELIMITADOR,
    agrupar_filas,
    es_numero_financiero,
    exportar,
    limpiar_pie_pagina,
    obtener_ruta_asset,
    reorganizar_excel,
    texto_delimitado_a_excel,
)

COLUMNAS_PN = ["FECHA", "DESCRIPCIÓN", "REFERENCIA", "VALOR"]
COLUMNAS_SOC = ["FECHA", "DESCRIPCIÓN", "SUCURSAL/CANAL", "REFERENCIA 1", "REFERENCIA 2", "DOCUMENTO", "VALOR"]
COLUMNAS_EXTRACTO = ["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO.", "VALOR", "SALDO"]

_FECHA_SOC = re.compile(r"^\b(?:\d{4}/\d{2}/\d{2}|\d{1,2}/\d{2}(?:/\d{2,4})?)\b")
_FECHA_PN = re.compile(r"^\d{1,2}\s+(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\.?(?:\s+\d{4})?", re.I)
_FECHA_EXTRACTO = re.compile(r"^\b\d{1,2}/\d{2}(?:/\d{2,4})?\b")


def extraer_lineas_movimientos_soc(pdf_path):
    lineas = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            anc = page.width
            limites = [(anc * 0.35, "DESCRIPCIÓN"), (anc * 0.50, "SUCURSAL/CANAL"), (anc * 0.63, "REFERENCIA 1"),
                       (anc * 0.73, "REFERENCIA 2"), (anc * 0.85, "DOCUMENTO")]
            transacciones, tx = [], None
            for fila in agrupar_filas(words, 3.0):
                fila = limpiar_pie_pagina(fila)
                if not fila:
                    continue
                primera = fila[0]["text"].strip()
                if _FECHA_SOC.match(primera) and fila[0]["x0"] < anc * 0.20:
                    if tx:
                        transacciones.append(tx)
                    tx = {"FECHA": primera, **{c: [] for _, c in limites}, "VALOR": ""}
                    fila = fila[1:]
                elif not tx:
                    continue
                for w in fila:
                    x = (w["x0"] + w["x1"]) / 2.0
                    col = next((c for lim, c in limites if x < lim), None)
                    if col:
                        tx[col].append(w["text"])
                    elif es_numero_financiero(w["text"]):
                        tx["VALOR"] = w["text"]
            if tx:
                transacciones.append(tx)
            for tx in transacciones:
                partes = [" ".join(tx[c]).strip() for _, c in limites]
                if "DESCRIPC" in partes[0].upper() or "REFERENCIA" in partes[2].upper():
                    continue
                lineas.append(DELIMITADOR.join([tx["FECHA"], *partes, tx["VALOR"]]))
    return lineas


def _es_dato_ip(texto):
    t = texto.strip().lower()
    return "direcci" in t or t == "ip" or "181.137" in t


def _monto_pn(cadena):
    num = re.sub(r"[^\d.,]", "", cadena)
    if "," in num and "." in num:
        num = num.replace(".", "").replace(",", ".")
    elif "," in num:
        num = num.replace(",", ".")
    m = re.search(r"\d+(?:\.\d+)?", num)
    if not m:
        return "0"
    return f"-{m.group(0)}" if "-" in cadena else m.group(0)


def extraer_lineas_movimientos_pn(pdf_path):
    lineas = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            anc = page.width
            x_desc, x_ref = anc * 0.50, anc * 0.76
            transacciones, tx = [], None
            for fila in agrupar_filas(words, 3.0):
                fila = [w for w in limpiar_pie_pagina(fila) if not _es_dato_ip(w["text"])]
                if not fila:
                    continue
                fecha = _FECHA_PN.match(" ".join(w["text"] for w in fila[:4]))
                if fecha and fila[0]["x0"] < anc * 0.25:
                    if tx:
                        transacciones.append(tx)
                    fecha_str = fecha.group(0).strip()
                    n = len(fecha_str.split())
                    if n < len(fila) and re.match(r"^\d{4}$", fila[n]["text"].strip()):
                        fecha_str = f"{' '.join(fecha_str.split()[:2])} {fila[n]['text'].strip()}"
                        n += 1
                    tx = {"FECHA": fecha_str, "DESCRIPCIÓN": [], "REFERENCIA": [], "VALOR_RAW": []}
                    fila = fila[n:]
                elif not tx:
                    continue
                for w in fila:
                    txt = w["text"].strip()
                    if not txt:
                        continue
                    x = (w["x0"] + w["x1"]) / 2.0
                    tx["DESCRIPCIÓN" if x < x_desc else "REFERENCIA" if x < x_ref else "VALOR_RAW"].append(txt)
            if tx:
                transacciones.append(tx)
            for tx in transacciones:
                desc = " ".join(tx["DESCRIPCIÓN"]).strip()
                if "DESCRIPCI" in desc.upper():
                    continue
                ref = " ".join(tx["REFERENCIA"]).strip()
                lineas.append(DELIMITADOR.join([tx["FECHA"], desc, ref, _monto_pn("".join(tx["VALOR_RAW"]).strip())]))
    return lineas


def extraer_lineas_extractos(pdf_path):
    lineas = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue
            anc = page.width
            for fila in agrupar_filas(words, 2.5):
                fila = limpiar_pie_pagina(fila)
                if not fila or fila[0]["x0"] >= anc * 0.20 or not _FECHA_EXTRACTO.match(fila[0]["text"].strip()):
                    continue
                resto = fila[1:]
                if not resto:
                    continue
                saldo = resto.pop()["text"] if es_numero_financiero(resto[-1]["text"]) else ""
                valor = resto.pop()["text"] if resto and es_numero_financiero(resto[-1]["text"]) else ""
                desc, suc, dcto = [], [], []
                for w in resto:
                    x = (w["x0"] + w["x1"]) / 2.0
                    (desc if x < anc * 0.40 else suc if x < anc * 0.65 else dcto).append(w["text"])
                descripcion = " ".join(desc).strip()
                if "DESCRIPC" in descripcion.upper() or "RESUMEN" in descripcion.upper():
                    continue
                lineas.append(DELIMITADOR.join([
                    fila[0]["text"].strip(), descripcion, " ".join(suc).strip(), " ".join(dcto).strip(), valor, saldo,
                ]))
    return lineas


def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None, progreso_callback=None):
    nombre = re.sub(r"[^A-Z]", "", os.path.splitext(os.path.basename(pdf_path))[0].upper())
    es_movimiento = "MOVIMIENTO" in nombre or "MOVIMENTO" in nombre

    if es_movimiento and "NATURAL" in nombre:
        columnas, lineas = COLUMNAS_PN, extraer_lineas_movimientos_pn(pdf_path)
    elif es_movimiento:
        columnas, lineas = COLUMNAS_SOC, extraer_lineas_movimientos_soc(pdf_path)
    elif "EXTRACTO" in nombre:
        columnas, lineas = COLUMNAS_EXTRACTO, extraer_lineas_extractos(pdf_path)
    else:
        raise ValueError(
            "El nombre del archivo no coincide con un prefijo válido "
            "('MovimientosPNatural', 'MovimientosSOC' o 'Extracto')."
        )

    if not lineas:
        return None
    return exportar(lineas, columnas, pdf_path, output_excel_path)