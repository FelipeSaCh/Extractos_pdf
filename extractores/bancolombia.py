import os
import re

from extractores import Extraccion
from infraestructura.lectura_pdf import agrupar_filas, centro_x, es_numero_financiero, leer_paginas, limpiar_pie_pagina
from nucleo.constantes import DELIMITADOR
from nucleo.numeros import a_float

COLUMNAS_PN = ["FECHA", "DESCRIPCIÓN", "REFERENCIA", "VALOR"]
COLUMNAS_SOC = ["FECHA", "DESCRIPCIÓN", "SUCURSAL/CANAL", "REFERENCIA 1", "REFERENCIA 2", "DOCUMENTO", "VALOR"]
COLUMNAS_EXTRACTO = ["FECHA", "DESCRIPCIÓN", "SUCURSAL", "DCTO.", "VALOR", "SALDO"]

_FECHA_SOC = re.compile(r"^\b(?:\d{4}/\d{2}/\d{2}|\d{1,2}/\d{2}(?:/\d{2,4})?)\b")
_FECHA_PN = re.compile(r"^\d{1,2}\s+(?:ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\.?(?:\s+\d{4})?", re.I)
_FECHA_EXTRACTO = re.compile(r"^\b\d{1,2}/\d{2}(?:/\d{2,4})?\b")


def _pagina_soc(page):
    words = page.extract_words()
    if not words:
        return []
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
            col = next((c for lim, c in limites if centro_x(w) < lim), None)
            if col:
                tx[col].append(w["text"])
            elif es_numero_financiero(w["text"]):
                tx["VALOR"] = w["text"]
    if tx:
        transacciones.append(tx)
    lineas = []
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
    return f"{a_float(cadena):.2f}"


def _pagina_pn(page):
    words = page.extract_words()
    if not words:
        return []
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
            if txt:
                x = centro_x(w)
                tx["DESCRIPCIÓN" if x < x_desc else "REFERENCIA" if x < x_ref else "VALOR_RAW"].append(txt)
    if tx:
        transacciones.append(tx)
    lineas = []
    for tx in transacciones:
        desc = " ".join(tx["DESCRIPCIÓN"]).strip()
        if "DESCRIPCI" in desc.upper():
            continue
        ref = " ".join(tx["REFERENCIA"]).strip()
        lineas.append(DELIMITADOR.join([tx["FECHA"], desc, ref, _monto_pn("".join(tx["VALOR_RAW"]).strip())]))
    return lineas


def _pagina_extracto(page):
    words = page.extract_words()
    if not words:
        return []
    anc = page.width
    lineas = []
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
            x = centro_x(w)
            (desc if x < anc * 0.40 else suc if x < anc * 0.65 else dcto).append(w["text"])
        descripcion = " ".join(desc).strip()
        if "DESCRIPC" in descripcion.upper() or "RESUMEN" in descripcion.upper():
            continue
        lineas.append(DELIMITADOR.join([
            fila[0]["text"].strip(), descripcion, " ".join(suc).strip(), " ".join(dcto).strip(), valor, saldo,
        ]))
    return lineas


TIPOS = {
    "MovimientosPNatural": (COLUMNAS_PN, _pagina_pn),
    "MovimientoSOC": (COLUMNAS_SOC, _pagina_soc),
    "Extracto": (COLUMNAS_EXTRACTO, _pagina_extracto),
}


def tipo_por_nombre(pdf_path):
    nombre = re.sub(r"[^A-Z]", "", os.path.splitext(os.path.basename(pdf_path))[0].upper())
    es_movimiento = "MOVIMIENTO" in nombre or "MOVIMENTO" in nombre
    if es_movimiento and "NATURAL" in nombre:
        return "MovimientosPNatural"
    if es_movimiento:
        return "MovimientoSOC"
    if "EXTRACTO" in nombre:
        return "Extracto"
    raise ValueError(
        "El nombre del archivo no coincide con un prefijo válido "
        "('MovimientosPNatural', 'MovimientosSOC' o 'Extracto')."
    )


def extraer(pdf_path, progreso=None, tipo_documento=None, **_):
    columnas, pagina = TIPOS[tipo_documento if tipo_documento in TIPOS else tipo_por_nombre(pdf_path)]
    lineas, cuenta = leer_paginas(pdf_path, pagina, progreso)
    return Extraccion(columnas, lineas, cuenta)