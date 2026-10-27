import re

import pdfplumber

from comun_extractos import DELIMITADOR, agrupar_filas, exportar

TIPO_COMUN = "COMUN"
TIPO_DAMAS = "DAMAS"
TIPOS_CUENTA = {
    TIPO_COMUN: "Cuenta de ahorros común (Valor / Clase de Movimiento)",
    TIPO_DAMAS: "Cuenta de ahorros DAMAS (Débito / Crédito)",
}
COLUMNAS = ["FECHA", "COD TRANS", "DESCRIPCIÓN", "CIUDAD", "OFICINA/CANAL", "DOCUMENTO", "VALOR", "SALDO"]

FRASES_PIE_PAGINA = [
    "LA IMPORTANCIA DEL AHORRO",
    "RESUMEN MES",
    "RECONOCERÁN A PARTIR DE SALDOS",
    "ESTE PRODUCTO CUENTA CON SEGURO",
    "CUALQUIER DIFERENCIA CON EL SALDO",
    "REVISORÍA",
    "DEFENSOR DEL CONSUMIDOR",
    "TELÉFONO:",
    "CORREO ELECTRÓNICO",
    "PARA MAYOR INFORMACIÓN EN",
    "DAVIVIENDA A PARTIR DEL",
]

_PATRON_MONTO = re.compile(r"^\$?-?[\d,]+\.\d{2}$")
_PATRON_ENTERO = re.compile(r"^\d+$")
_PATRON_DIA_MES = re.compile(r"^\d{1,2}$")
_PATRON_SALDO = re.compile(r"^-?[\d.]+$")
_PATRON_ANO = re.compile(
    r"(?i)(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s*/\s*(20\d{2})"
)
_ETIQUETAS_SALDO = {("SALDO", "ANTERIOR"): "anterior", ("NUEVO", "SALDO"): "nuevo"}


def _tipo_desde_texto(texto):
    texto = (texto or "").upper()
    if "CLASE DE MOVIMIENTO" in texto:
        return TIPO_COMUN
    if re.search(r"DESCRIPCI[OÓ]N.*D[EÉ]BITO.*CR[EÉ]DITO", texto):
        return TIPO_DAMAS
    return None


def detectar_tipo_cuenta(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return _tipo_desde_texto(pdf.pages[0].extract_text())


def elegir_tipo_cuenta(parent, pdf_path=None):
    import tkinter as tk
    from tkinter import ttk

    try:
        sugerido = detectar_tipo_cuenta(pdf_path) if pdf_path else None
    except Exception:
        sugerido = None
    var = tk.StringVar(value=sugerido or TIPO_COMUN)
    res = {}

    ven = tk.Toplevel(parent)
    ven.title("Tipo de cuenta Davivienda")
    ven.transient(parent)
    ven.resizable(False, False)
    frm = ttk.Frame(ven, padding=16)
    frm.pack(fill=tk.BOTH, expand=True)
    ttk.Label(frm, text="Seleccione el tipo de extracto:", font=("Segoe UI", 10)).pack(anchor="w", pady=(0, 8))
    for clave, texto in TIPOS_CUENTA.items():
        etiqueta = f"{texto}  (detectado)" if clave == sugerido else texto
        ttk.Radiobutton(frm, text=etiqueta, variable=var, value=clave).pack(anchor="w", pady=2)

    def aceptar(_=None):
        res["tipo"] = var.get()
        ven.destroy()

    botones = ttk.Frame(frm)
    botones.pack(fill=tk.X, pady=(12, 0))
    ttk.Button(botones, text="Cancelar", command=ven.destroy).pack(side=tk.RIGHT)
    ttk.Button(botones, text="Aceptar", command=aceptar).pack(side=tk.RIGHT, padx=6)
    ven.bind("<Return>", aceptar)
    ven.bind("<Escape>", lambda e: ven.destroy())
    ven.grab_set()
    parent.wait_window(ven)
    return res.get("tipo")


def _a_float(txt):
    return float(txt.replace("$", "").replace(",", ""))


def _fecha(dia, mes, year):
    base = f"{dia.zfill(2)}/{mes.zfill(2)}"
    return f"{base}/{year}" if year else base


def _saldos(words):
    saldos = {"anterior": None, "nuevo": None}
    for i in range(len(words) - 1):
        clave = _ETIQUETAS_SALDO.get((words[i]["text"].upper(), words[i + 1]["text"].upper()))
        if not clave:
            continue
        for w in words[i + 2:i + 6]:
            txt = w["text"].replace("$", "").replace(",", "")
            if _PATRON_SALDO.match(txt):
                saldos[clave] = float(txt)
                break
    return saldos["anterior"], saldos["nuevo"]


def _parsear_fila_damas(palabras):
    tokens = [w for w in palabras if w["text"].strip() and w["text"].strip() != "$"]
    textos = [w["text"].strip() for w in tokens]
    if len(textos) < 6 or not (_PATRON_ENTERO.match(textos[0]) and _PATRON_ENTERO.match(textos[1])):
        return None

    montos = []
    while textos and len(montos) < 2 and _PATRON_MONTO.match(textos[-1]):
        montos.insert(0, _a_float(textos.pop()))
        tokens.pop()
    if len(montos) < 2:
        return None

    doc = ""
    if len(textos) > 3 and _PATRON_ENTERO.match(textos[-1]):
        doc = textos.pop()
        tokens.pop()

    inicio_desc, oficina = 2, ""
    if len(textos) > 2 and _PATRON_ENTERO.match(textos[2]):
        oficina, inicio_desc = textos[2], 3

    return {
        "dia": textos[0],
        "mes": textos[1],
        "oficina": oficina,
        "descripcion": textos[inicio_desc:],
        "x_desc": tokens[inicio_desc]["x0"] if len(tokens) > inicio_desc else None,
        "doc": doc,
        "valor": montos[1] - montos[0],
    }


def _es_pie(texto):
    return any(frase in texto for frase in FRASES_PIE_PAGINA)


def _filas_damas(filas, anc, year, salida):
    tx = None
    for palabras in filas:
        texto = " ".join(w["text"] for w in palabras).upper()
        if _es_pie(texto):
            break
        if "DESCRIPCI" in texto and ("DÉBITO" in texto or "DEBITO" in texto):
            continue
        fila = _parsear_fila_damas(palabras) if palabras[0]["x0"] < anc * 0.15 else None
        if fila:
            if tx:
                _guardar_tx(tx, salida)
            tx = {
                "FECHA": _fecha(fila["dia"], fila["mes"], year),
                "VALOR": [f"{fila['valor']:.2f}"],
                "DOCUMENTO": [fila["doc"]],
                "DESCRIPCIÓN": fila["descripcion"],
                "OFICINA": [fila["oficina"]],
                "_x_desc": fila["x_desc"],
            }
        elif (
            tx
            and tx["_x_desc"] is not None
            and palabras[0]["x0"] >= tx["_x_desc"] - 2
            and not any(_PATRON_MONTO.match(w["text"].strip()) for w in palabras)
        ):
            tx["DESCRIPCIÓN"].extend(w["text"] for w in palabras)
    if tx:
        _guardar_tx(tx, salida)


def _filas_comun(filas, anc, year, salida):
    tx = None
    for palabras in filas:
        texto = " ".join(w["text"] for w in palabras).upper()
        if _es_pie(texto):
            break
        if "FECHA" in texto and "VALOR" in texto and "CLASE DE MOVIMIENTO" in texto:
            continue
        dia, mes = (palabras[0]["text"].strip(), palabras[1]["text"].strip()) if len(palabras) >= 2 else ("", "")
        if _PATRON_DIA_MES.match(dia) and _PATRON_DIA_MES.match(mes) and palabras[0]["x0"] < anc * 0.10:
            if tx:
                _guardar_tx(tx, salida)
            tx = {"FECHA": _fecha(dia, mes, year), "VALOR": [], "DOCUMENTO": [], "DESCRIPCIÓN": [], "OFICINA": []}
            for w in palabras[2:]:
                x = (w["x0"] + w["x1"]) / 2.0
                clave = ("VALOR" if x < anc * 0.28 else "DOCUMENTO" if x < anc * 0.35
                         else "DESCRIPCIÓN" if x < anc * 0.72 else "OFICINA")
                tx[clave].append(w["text"])
        elif tx and palabras[0]["x0"] >= anc * 0.35:
            for w in palabras:
                x = (w["x0"] + w["x1"]) / 2.0
                tx["DESCRIPCIÓN" if x < anc * 0.72 else "OFICINA"].append(w["text"])
    if tx:
        _guardar_tx(tx, salida)


def extraer_lineas_davivienda(pdf_path, progreso_callback=None, tipo_cuenta=None):
    lineas = []
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        primera = pdf.pages[0]
        texto_pag1 = primera.extract_text()
        match_ano = _PATRON_ANO.search(texto_pag1 or "")
        year = match_ano.group(1) if match_ano else None

        detectado = _tipo_desde_texto(texto_pag1)
        if tipo_cuenta and detectado and tipo_cuenta != detectado:
            raise ValueError(
                f"El extracto corresponde a '{TIPOS_CUENTA[detectado]}', "
                f"pero se seleccionó '{TIPOS_CUENTA[tipo_cuenta]}'."
            )
        procesar = _filas_damas if (tipo_cuenta or detectado or TIPO_COMUN) == TIPO_DAMAS else _filas_comun
        saldo_anterior, nuevo_saldo = _saldos(primera.extract_words())

        for n_pag, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()
            if words:
                procesar(agrupar_filas(words, 2.0), page.width, year, lineas)
            if progreso_callback:
                progreso_callback(n_pag, total)

    return lineas, saldo_anterior, nuevo_saldo


def _guardar_tx(tx, salida):
    desc = " ".join(tx["DESCRIPCIÓN"]).strip()
    oficina = " ".join(tx["OFICINA"]).strip()
    doc = " ".join(tx["DOCUMENTO"]).strip()
    bruto = " ".join(tx["VALOR"]).replace("$", "").replace(",", "").replace(" ", "").strip()

    if bruto.endswith("-"):
        valor = f"-{bruto[:-1]}"
    elif bruto.endswith("+"):
        valor = bruto[:-1]
    else:
        valor = bruto or "0.0"

    salida.append(DELIMITADOR.join([tx["FECHA"], "", desc, "", oficina, doc, valor, ""]))


def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None, progreso_callback=None, tipo_cuenta=None):
    lineas, saldo_ant, saldo_nue = extraer_lineas_davivienda(pdf_path, progreso_callback, tipo_cuenta)
    if not lineas:
        raise ValueError("No se encontraron movimientos del extracto o el formato es irreconocible.")
    return exportar(lineas, COLUMNAS, pdf_path, output_excel_path, saldo_ant, saldo_nue)