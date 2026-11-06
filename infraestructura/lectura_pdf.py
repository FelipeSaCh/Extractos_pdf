import re

import pdfplumber

from nucleo.texto import sin_tildes

_PATRON_FINANCIERO = re.compile(r"^-?[\d.,]+$")
_PATRON_PAG = re.compile(r"(?i)^p[áa]g(?:ina)?\.?$")
_SECUENCIA_PAG = (re.compile(r"^\d+$"), re.compile(r"(?i)^de\.?$"), re.compile(r"^\d+$"))

FRASES_CUENTA = (
    "NUMERO DE CUENTA", "NUMERO DE LA CUENTA", "NO. DE CUENTA", "NO DE CUENTA", "NRO. CUENTA", "NRO CUENTA",
    "CUENTA NO.", "CUENTA NO", "CUENTA NRO.", "CUENTA NRO", "CUENTA N°", "CUENTA #",
    "CUENTA DE AHORROS", "CUENTA AHORROS", "CUENTA CORRIENTE", "CTA. AHORROS", "CTA AHORROS",
    "CTA. CORRIENTE", "CTA CORRIENTE", "CTA.", "CTA", "NUMERO DE PRODUCTO", "PRODUCTO", "NÚMERO",
)
_NUMERO_CUENTA = r"[^\d\n]{0,25}\n?[^\d\n]{0,10}(\d[\d -]{4,24}\d)"
PAGINAS_CUENTA = 2


def _patron_frase(frase):
    palabras = [re.escape(p) for p in sin_tildes(frase).split()]
    return re.compile(r"(?<![A-Z])" + r"\s*".join(palabras) + r"(?![A-Z])" + _NUMERO_CUENTA)


_PATRONES_CUENTA = [_patron_frase(f) for f in FRASES_CUENTA]


def buscar_cuenta_bancaria(texto):
    texto = sin_tildes(texto or "")
    for patron in _PATRONES_CUENTA:
        for m in patron.finditer(texto):
            numero = re.sub(r"\s", "", m[1])
            if sum(c.isdigit() for c in numero) >= 6:
                return numero
    return ""


def es_numero_financiero(texto):
    return bool(texto) and bool(_PATRON_FINANCIERO.match(texto.strip().replace("$", "").replace(" ", "")))


def limpiar_pie_pagina(words_fila, frases=()):
    if frases:
        linea = " ".join(w["text"] for w in words_fila).upper()
        if any(f in linea for f in frases):
            return []
    textos = [w["text"].strip() for w in words_fila]
    quitar, i = set(), 0
    while i < len(textos):
        if not _PATRON_PAG.match(textos[i]):
            i += 1
            continue
        j = i + 1
        for patron in _SECUENCIA_PAG:
            if j >= len(textos) or not patron.match(textos[j]):
                break
            j += 1
        quitar.update(range(i, j))
        i = j
    return [w for k, w in enumerate(words_fila) if k not in quitar] if quitar else words_fila


def agrupar_filas(words, paso):
    filas = {}
    for w in words:
        filas.setdefault(round(w["top"] / paso) * paso, []).append(w)
    return [sorted(filas[y], key=lambda w: w["x0"]) for y in sorted(filas)]


def centro_x(w):
    return (w["x0"] + w["x1"]) / 2.0


def leer_paginas(pdf_path, procesar_pagina, progreso=None):
    """Abre el PDF una sola vez: aplica procesar_pagina(page) -> list a cada página,
    reporta progreso(n, total) y detecta la cuenta bancaria en las primeras páginas."""
    lineas, textos = [], []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        for n, page in enumerate(pdf.pages, start=1):
            if n <= PAGINAS_CUENTA:
                textos.append(page.extract_text() or "")
            lineas += procesar_pagina(page)
            if progreso:
                progreso(n, total)
    return lineas, buscar_cuenta_bancaria("\n".join(textos))


def texto_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(p.extract_text() or "" for p in pdf.pages)


def texto_y_tablas(pdf_path):
    """Texto (ignorando caracteres rotados) y tablas de todas las páginas."""
    textos, bloques = [], []
    with pdfplumber.open(pdf_path) as pdf:
        for pagina in pdf.pages:
            p = pagina.filter(lambda o: o.get("object_type") != "char" or o.get("upright", True))
            textos.append(p.extract_text() or "")
            bloques += [[[str(c or "").strip() for c in f] for f in t] for t in p.extract_tables()]
    return "\n".join(textos), bloques