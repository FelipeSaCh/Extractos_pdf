import os
import re
import pandas as pd
import pdfplumber
from script import reorganizar_excel

DELIMITADOR = "|||"

TIPO_COMUN = "COMUN"
TIPO_DAMAS = "DAMAS"
TIPOS_CUENTA = {
    TIPO_COMUN: "Cuenta de ahorros común (Valor / Clase de Movimiento)",
    TIPO_DAMAS: "Cuenta de ahorros DAMAS (Débito / Crédito)",
}

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
    "DAVIVIENDA A PARTIR DEL"
]

_PATRON_MONTO = re.compile(r"^\$?-?[\d,]+\.\d{2}$")
_PATRON_ENTERO = re.compile(r"^\d+$")


def texto_delimitado_a_excel(lineas_texto, columnas, output_excel_path):
    """Convierte las líneas delimitadas a un DataFrame y lo exporta a Excel."""
    registros = []
    for linea in lineas_texto:
        partes = linea.split(DELIMITADOR)
        if len(partes) == len(columnas):
            registros.append(dict(zip(columnas, partes)))

    df = pd.DataFrame(registros, columns=columnas)
    df.to_excel(output_excel_path, index=False)


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
    """Diálogo para que el usuario seleccione el tipo de cuenta Davivienda. Retorna la clave o None."""
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


def _parsear_fila_damas(palabras):
    """Fila DAMAS: Día Mes Oficina Descripción... Doc Débito Crédito."""
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

    inicio_desc = 2
    oficina = ""
    if len(textos) > 2 and _PATRON_ENTERO.match(textos[2]):
        oficina = textos[2]
        inicio_desc = 3

    return {
        "dia": textos[0],
        "mes": textos[1],
        "oficina": oficina,
        "descripcion": textos[inicio_desc:],
        "x_desc": tokens[inicio_desc]["x0"] if len(tokens) > inicio_desc else None,
        "doc": doc,
        "valor": montos[1] - montos[0],
    }


def _fecha(dia, mes, year):
    base = f"{dia.zfill(2)}/{mes.zfill(2)}"
    return f"{base}/{year}" if year else base


def extraer_lineas_davivienda(pdf_path, progreso_callback=None, tipo_cuenta=None):
    """Procesa el PDF de Davivienda en formato común o DAMAS."""
    lineas_delimitadas = []
    year = None
    saldo_anterior = None
    nuevo_saldo = None

    with pdfplumber.open(pdf_path) as pdf:
        total_paginas = len(pdf.pages)

        # --- 1. Extracción de Año, Tipo y Saldos en la Primera Página ---
        primer_pag = pdf.pages[0]
        texto_pag1 = primer_pag.extract_text()

        if texto_pag1:
            match_year = re.search(r'(?i)(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s*/\s*(20\d{2})', texto_pag1)
            if match_year:
                year = match_year.group(1)

        detectado = _tipo_desde_texto(texto_pag1)
        if tipo_cuenta and detectado and tipo_cuenta != detectado:
            raise ValueError(
                f"El extracto corresponde a '{TIPOS_CUENTA[detectado]}', "
                f"pero se seleccionó '{TIPOS_CUENTA[tipo_cuenta]}'."
            )
        tipo = tipo_cuenta or detectado or TIPO_COMUN

        words_pag1 = primer_pag.extract_words()
        for i, w in enumerate(words_pag1):
            txt = w['text'].upper()

            if txt == "SALDO" and i + 1 < len(words_pag1) and words_pag1[i+1]['text'].upper() == "ANTERIOR":
                for j in range(i+2, min(i+6, len(words_pag1))):
                    val_txt = words_pag1[j]['text'].replace('$', '').replace(',', '')
                    if re.match(r'^-?[\d\.]+$', val_txt):
                        saldo_anterior = float(val_txt)
                        break

            elif txt == "NUEVO" and i + 1 < len(words_pag1) and words_pag1[i+1]['text'].upper() == "SALDO":
                for j in range(i+2, min(i+6, len(words_pag1))):
                    val_txt = words_pag1[j]['text'].replace('$', '').replace(',', '')
                    if re.match(r'^-?[\d\.]+$', val_txt):
                        nuevo_saldo = float(val_txt)
                        break
        # ---------------------------------------------------------

        for indice_pagina, page in enumerate(pdf.pages, start=1):
            anc = page.width
            words = page.extract_words()

            if not words:
                if progreso_callback: progreso_callback(indice_pagina, total_paginas)
                continue

            lineas = {}
            for w in words:
                top_key = round(w['top'] / 2.0) * 2.0
                lineas.setdefault(top_key, []).append(w)

            tx_actual = None

            for y in sorted(lineas):
                palabras_linea = sorted(lineas[y], key=lambda w: w['x0'])
                texto_linea = " ".join([w['text'] for w in palabras_linea]).upper()

                if any(frase in texto_linea for frase in FRASES_PIE_PAGINA):
                    break

                if tipo == TIPO_DAMAS:
                    if "DESCRIPCI" in texto_linea and ("DÉBITO" in texto_linea or "DEBITO" in texto_linea):
                        continue

                    fila = _parsear_fila_damas(palabras_linea) if palabras_linea[0]['x0'] < anc * 0.15 else None
                    if fila:
                        if tx_actual:
                            _guardar_tx(tx_actual, lineas_delimitadas)
                        tx_actual = {
                            "FECHA": _fecha(fila["dia"], fila["mes"], year),
                            "VALOR": [f"{fila['valor']:.2f}"],
                            "DOCUMENTO": [fila["doc"]],
                            "DESCRIPCIÓN": fila["descripcion"],
                            "OFICINA": [fila["oficina"]],
                            "_x_desc": fila["x_desc"],
                        }
                    elif (
                        tx_actual
                        and tx_actual["_x_desc"] is not None
                        and palabras_linea[0]['x0'] >= tx_actual["_x_desc"] - 2
                        and not any(_PATRON_MONTO.match(w['text'].strip()) for w in palabras_linea)
                    ):
                        tx_actual["DESCRIPCIÓN"].extend(w['text'] for w in palabras_linea)
                    continue

                # --- Formato común ---
                if "FECHA" in texto_linea and "VALOR" in texto_linea and "CLASE DE MOVIMIENTO" in texto_linea:
                    continue

                es_nueva_tx = False
                if len(palabras_linea) >= 2:
                    dia_text = palabras_linea[0]['text'].strip()
                    mes_text = palabras_linea[1]['text'].strip()
                    if re.match(r'^\d{1,2}$', dia_text) and re.match(r'^\d{1,2}$', mes_text) and palabras_linea[0]['x0'] < anc * 0.10:
                        es_nueva_tx = True

                if es_nueva_tx:
                    if tx_actual:
                        _guardar_tx(tx_actual, lineas_delimitadas)

                    tx_actual = {
                        "FECHA": _fecha(dia_text, mes_text, year),
                        "VALOR": [],
                        "DOCUMENTO": [],
                        "DESCRIPCIÓN": [],
                        "OFICINA": []
                    }

                    for w in palabras_linea[2:]:
                        x_mid = (w['x0'] + w['x1']) / 2.0
                        if x_mid < anc * 0.28:
                            tx_actual["VALOR"].append(w['text'])
                        elif x_mid < anc * 0.35:
                            tx_actual["DOCUMENTO"].append(w['text'])
                        elif x_mid < anc * 0.72:
                            tx_actual["DESCRIPCIÓN"].append(w['text'])
                        else:
                            tx_actual["OFICINA"].append(w['text'])
                elif tx_actual and palabras_linea and anc * 0.35 <= palabras_linea[0]['x0']:
                    for w in palabras_linea:
                        x_mid = (w['x0'] + w['x1']) / 2.0
                        if x_mid < anc * 0.72:
                            tx_actual["DESCRIPCIÓN"].append(w['text'])
                        else:
                            tx_actual["OFICINA"].append(w['text'])

            if tx_actual:
                _guardar_tx(tx_actual, lineas_delimitadas)
                tx_actual = None

            if progreso_callback:
                progreso_callback(indice_pagina, total_paginas)

    return lineas_delimitadas, saldo_anterior, nuevo_saldo


def _guardar_tx(tx_dict, lista_salida):
    """Procesa el valor numérico moviendo el signo final al inicio y guarda el registro."""
    desc = " ".join(tx_dict["DESCRIPCIÓN"]).strip()
    oficina = " ".join(tx_dict["OFICINA"]).strip()
    doc = " ".join(tx_dict["DOCUMENTO"]).strip()

    valor_str_bruto = " ".join(tx_dict["VALOR"]).replace('$', '').replace(',', '').replace(' ', '').strip()

    valor_final = "0.0"
    if valor_str_bruto:
        if valor_str_bruto.endswith('-'):
            valor_final = f"-{valor_str_bruto[:-1]}"
        elif valor_str_bruto.endswith('+'):
            valor_final = valor_str_bruto[:-1]
        else:
            valor_final = valor_str_bruto

    registro = [
        tx_dict["FECHA"],
        "",         # COD TRANS
        desc,
        "",         # CIUDAD
        oficina,
        doc,
        valor_final,
        ""          # SALDO
    ]
    lista_salida.append(DELIMITADOR.join(registro))


def inyectar_saldos_en_resumen(excel_path, saldo_ant, saldo_nue):
    """Inserta los saldos del extracto en la hoja Resumen y la diferencia de conciliación."""
    if saldo_ant is None and saldo_nue is None:
        return

    xls = pd.ExcelFile(excel_path)
    df_datos = pd.read_excel(xls, sheet_name="Datos")
    df_resumen = pd.read_excel(xls, sheet_name="Resumen")
    df_conceptos = pd.read_excel(xls, sheet_name="Conceptos")

    nuevas_filas = []

    if saldo_ant is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ANTERIOR", "MONTO": saldo_ant})

    filas_limpias = [
        f for f in df_resumen.to_dict('records')
        if f["CONCEPTO"] not in ["SALDO ANTERIOR", "SALDO ACTUAL", "DIFERENCIA CONCILIACIÓN"]
    ]
    nuevas_filas.extend(filas_limpias)

    if saldo_nue is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ACTUAL", "MONTO": saldo_nue})

    if saldo_ant is not None and saldo_nue is not None:
        movimientos = sum(
            float(f["MONTO"]) for f in filas_limpias if f["CONCEPTO"] in ["TOTAL CARGOS", "TOTAL ABONOS"]
        )
        nuevas_filas.append({
            "CONCEPTO": "DIFERENCIA CONCILIACIÓN",
            "MONTO": round(saldo_ant + movimientos - saldo_nue, 2),
        })

    df_resumen_actualizado = pd.DataFrame(nuevas_filas)

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_datos.to_excel(writer, sheet_name="Datos", index=False)
        df_resumen_actualizado.to_excel(writer, sheet_name="Resumen", index=False)
        df_conceptos.to_excel(writer, sheet_name="Conceptos", index=False)

        workbook = writer.book
        FORMATO_MONEDA = '"$"#,##0.00;[Red]("$"#,##0.00);"-"'

        for sheet_name in workbook.sheetnames:
            worksheet = workbook[sheet_name]
            for col in worksheet.iter_cols(1, worksheet.max_column):
                header_val = str(col[0].value).upper() if col[0].value else ""
                if header_val in ["VALOR", "SALDO", "CARGOS", "ABONOS", "NETO", "MONTO"]:
                    for cell in col[1:]:
                        if cell.value is not None and isinstance(cell.value, (int, float)):
                            cell.number_format = FORMATO_MONEDA


def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None, progreso_callback=None, tipo_cuenta=None):
    """Función orquestadora principal para la importación desde app_gui."""
    columnas = [
        "FECHA",
        "COD TRANS",
        "DESCRIPCIÓN",
        "CIUDAD",
        "OFICINA/CANAL",
        "DOCUMENTO",
        "VALOR",
        "SALDO",
    ]

    lineas_plana, saldo_ant, saldo_nue = extraer_lineas_davivienda(pdf_path, progreso_callback, tipo_cuenta)

    if not lineas_plana:
        raise ValueError("No se encontraron movimientos del extracto o el formato es irreconocible.")

    if not output_excel_path:
        base_path, _ = os.path.splitext(pdf_path)
        output_excel_path = f"{base_path}_convertido.xlsx"

    base_txt_path, _ = os.path.splitext(output_excel_path)
    txt_path = f"{base_txt_path}_plano.txt"

    encabezado_txt = DELIMITADOR.join(columnas)
    contenido_txt = [encabezado_txt] + lineas_plana

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(contenido_txt))

    texto_delimitado_a_excel(lineas_plana, columnas, output_excel_path)
    reorganizar_excel(output_excel_path)
    inyectar_saldos_en_resumen(output_excel_path, saldo_ant, saldo_nue)

    return output_excel_path