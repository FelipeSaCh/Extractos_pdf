import re

from extractores import Extraccion
from infraestructura.lectura_pdf import centro_x, es_numero_financiero, leer_paginas
from nucleo.constantes import DELIMITADOR

COLUMNAS = ["FECHA", "COD TRANS", "DESCRIPCIÓN", "CIUDAD", "OFICINA/CANAL", "DOCUMENTO", "VALOR", "SALDO"]
PATRON_FECHA = re.compile(r"^\b\d{1,2}/\d{2}\b")
MARCAS_FIN = ("FIN MOVIMIENTOS", "NOTIFICAR", "REVISORÍA")
SIN_MOVIMIENTOS = "No se pudieron extraer movimientos del PDF."


def _movimientos_pagina(page):
    anc, alt = page.width, page.height
    words = [w for w in page.extract_words() if anc * 0.05 <= w["x0"] <= anc * 0.98]
    if not words:
        return []

    limites = [(anc * 0.16, "COD_TRANS"), (anc * 0.43, "DESCRIPCIÓN"), (anc * 0.53, "CIUDAD"),
               (anc * 0.67, "OFICINA"), (anc * 0.74, "DOCUMENTO")]
    x_fecha, x_valor = anc * 0.11, anc * 0.88

    y_max = next((l["top"] - 2 for l in page.extract_text_lines()
                  if any(m in l["text"].upper() for m in MARCAS_FIN)), alt)
    anclas = [w for w in sorted(words, key=lambda w: (w["top"], w["x0"]))
              if w["top"] < y_max and w["x0"] < x_fecha and PATRON_FECHA.match(w["text"].strip())]

    lineas = []
    for i, ancla in enumerate(anclas):
        y0 = ancla["top"] - 1.5
        y1 = anclas[i + 1]["top"] - 1.5 if i + 1 < len(anclas) else y_max
        bloque = sorted((w for w in words if y0 <= w["top"] < y1 and w["top"] < y_max),
                        key=lambda w: (round(w["top"] / 2.0) * 2.0, w["x0"]))
        tx = {c: [] for _, c in limites}
        valor = saldo = ""
        for w in bloque:
            txt = w["text"].strip()
            if not txt or set(txt) == {"-"} or w == ancla:
                continue
            x = centro_x(w)
            col = next((c for lim, c in limites if x < lim), None)
            if col:
                tx[col].append(txt)
            elif x < x_valor:
                if not valor and es_numero_financiero(txt):
                    valor = txt
            elif not saldo and es_numero_financiero(txt):
                saldo = txt

        if "DESCRIPCI" in " ".join(tx["DESCRIPCIÓN"]).upper():
            continue
        lineas.append(DELIMITADOR.join(
            [ancla["text"].strip(), *(" ".join(tx[c]).strip() for _, c in limites), valor, saldo]
        ))
    return lineas


def extraer(pdf_path, progreso=None, **_):
    return Extraccion(COLUMNAS, *leer_paginas(pdf_path, _movimientos_pagina, progreso))