import re

from extractores import Extraccion
from infraestructura.lectura_pdf import centro_x, es_numero_financiero, leer_paginas
from nucleo.constantes import DELIMITADOR

COLUMNAS = ["FECHA", "DESCRIPCIÓN", "VALOR", "SALDO"]
PATRON_FECHA = re.compile(r"^\d{2}/\d{2}/\d{4}$")
TEXTOS_FIN = ("DEFENSOR CONSUMIDOR", "NEQUI GENERAN", "PUEDES CONSULTAR", "VIGILADO", "SUPERINTENDENCIA")
ENCABEZADOS = {"FECHA", "DEL", "MOVIMIENTO", "DESCRIPCIÓN", "VALOR", "SALDO"}
SIN_MOVIMIENTOS = "No se pudieron extraer movimientos del PDF de Nequi."


def _movimientos_pagina(page):
    ancho, alto = page.width, page.height
    words = page.extract_words(x_tolerance=2, y_tolerance=3, keep_blank_chars=False)
    if not words:
        return []

    ordenadas = sorted(words, key=lambda w: (w["top"], w["x0"]))
    y_ini = next((w["top"] for w in ordenadas if w["text"].strip().upper().startswith("DESCRIPCI")), 0)
    y_fin = next((l["top"] - 3 for l in page.extract_text_lines()
                  if l["top"] > y_ini and any(t in l["text"].upper() for t in TEXTOS_FIN)), alto)
    tabla = [w for w in words if y_ini <= w["top"] < y_fin]

    x_fecha, x_desc, x_valor = ancho * 0.30, ancho * 0.65, ancho * 0.82
    anclas = [w for w in tabla if PATRON_FECHA.match(w["text"].strip()) and w["x0"] < x_fecha]

    lineas = []
    for i, ancla in enumerate(anclas):
        y0 = ancla["top"] - 2
        y1 = anclas[i + 1]["top"] - 2 if i + 1 < len(anclas) else y_fin
        bloque = sorted((w for w in tabla if y0 <= w["top"] < y1),
                        key=lambda w: (round(w["top"] / 2.0) * 2.0, w["x0"]))
        desc, valor, saldo = [], "", ""
        for w in bloque:
            txt = w["text"].strip()
            if not txt or w is ancla or txt.upper() in ENCABEZADOS:
                continue
            x = centro_x(w)
            if x < x_desc:
                if x < x_fecha and PATRON_FECHA.match(txt):
                    continue
                desc.append(txt)
            elif x < x_valor:
                if not valor and es_numero_financiero(txt):
                    valor = txt
            elif not saldo and es_numero_financiero(txt):
                saldo = txt

        descripcion = " ".join(desc).strip()
        if not (descripcion and valor and saldo):
            continue
        if "DESCRIPCI" in descripcion.upper() or "FECHA DEL MOVIMIENTO" in descripcion.upper():
            continue
        lineas.append(DELIMITADOR.join([ancla["text"].strip(), descripcion, valor, saldo]))
    return lineas


def extraer(pdf_path, progreso=None, **_):
    return Extraccion(COLUMNAS, *leer_paginas(pdf_path, _movimientos_pagina, progreso))