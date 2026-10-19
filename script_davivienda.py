import os
import re
import pandas as pd
import pdfplumber
from script import reorganizar_excel

DELIMITADOR = "|||"

def texto_delimitado_a_excel(lineas_texto, columnas, output_excel_path):
    """Convierte las líneas delimitadas a un DataFrame y lo exporta a Excel."""
    registros = []
    for linea in lineas_texto:
        partes = linea.split(DELIMITADOR)
        if len(partes) == len(columnas):
            registros.append(dict(zip(columnas, partes)))

    df = pd.DataFrame(registros, columns=columnas)
    df.to_excel(output_excel_path, index=False)

def extraer_lineas_davivienda(pdf_path, progreso_callback=None):
    """Procesa el PDF reconociendo el nuevo formato de tabla de Davivienda."""
    lineas_delimitadas = []
    year = None
    saldo_anterior = None
    nuevo_saldo = None

    with pdfplumber.open(pdf_path) as pdf:
        total_paginas = len(pdf.pages)
        
        # --- 1. Extracción de Año y Saldos en la Primera Página ---
        primer_pag = pdf.pages[0]
        texto_pag1 = primer_pag.extract_text()
        
        if texto_pag1:
            match_year = re.search(r'(?i)(?:enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)\s*/\s*(20\d{2})', texto_pag1)
            if match_year:
                year = match_year.group(1)

        # Buscar "Saldo Anterior" y "Nuevo Saldo" 
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

            # Agrupar palabras en "Líneas" visuales
            lineas = {}
            for w in words:
                top_key = round(w['top'] / 2.0) * 2.0
                if top_key not in lineas:
                    lineas[top_key] = []
                lineas[top_key].append(w)

            y_keys = sorted(list(lineas.keys()))
            tx_actual = None

            for y in y_keys:
                palabras_linea = sorted(lineas[y], key=lambda w: w['x0'])
                texto_linea = " ".join([w['text'] for w in palabras_linea]).upper()

                # --- Exclusiones de pie de página ---
                frases_pie_pagina = [
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
                
                if any(frase in texto_linea for frase in frases_pie_pagina):
                    break 
                
                # Ignorar encabezados de la nueva tabla
                if "FECHA" in texto_linea and "VALOR" in texto_linea and "CLASE DE MOVIMIENTO" in texto_linea:
                    continue
                
                es_nueva_tx = False
                if len(palabras_linea) >= 2:
                    dia_text = palabras_linea[0]['text'].strip()
                    mes_text = palabras_linea[1]['text'].strip()
                    # Detectar inicio de transacción (Día y Mes al extremo izquierdo)
                    if re.match(r'^\d{1,2}$', dia_text) and re.match(r'^\d{1,2}$', mes_text) and palabras_linea[0]['x0'] < anc * 0.10:
                        es_nueva_tx = True

                if es_nueva_tx:
                    if tx_actual:
                        _guardar_tx(tx_actual, lineas_delimitadas)

                    if year:
                        fecha_formateada = f"{dia_text.zfill(2)}/{mes_text.zfill(2)}/{year}"
                    else:
                        fecha_formateada = f"{dia_text.zfill(2)}/{mes_text.zfill(2)}"

                    tx_actual = {
                        "FECHA": fecha_formateada,
                        "VALOR": [],
                        "DOCUMENTO": [],
                        "DESCRIPCIÓN": [],
                        "OFICINA": []
                    }

                    # Clasificación por coordenadas para el nuevo formato
                    for w in palabras_linea[2:]:
                        x_mid = (w['x0'] + w['x1']) / 2.0
                        if x_mid < anc * 0.28:      # Columna Valor
                            tx_actual["VALOR"].append(w['text'])
                        elif x_mid < anc * 0.35:    # Columna Doc.
                            tx_actual["DOCUMENTO"].append(w['text'])
                        elif x_mid < anc * 0.72:    # Columna Clase de Movimiento
                            tx_actual["DESCRIPCIÓN"].append(w['text'])
                        else:                       # Columna Oficina
                            tx_actual["OFICINA"].append(w['text'])
                else:
                    # Lógica para líneas adicionales (ej. multilínea en "Clase de Movimiento" u "Oficina")
                    if tx_actual and palabras_linea:
                        x_primera = palabras_linea[0]['x0']
                        # Si la primera palabra cae dentro del área de Descripción u Oficina
                        if anc * 0.35 <= x_primera:
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

    valor_str_bruto = " ".join(tx_dict["VALOR"]).replace('$', '').replace(',', '').strip()
    
    # Manejar el signo al final del string (ej: "1359367.00-" o "16970338.00+")
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
    """Inserta los saldos encontrados directamente en la hoja Resumen, limpiando duplicados falsos."""
    if saldo_ant is None and saldo_nue is None:
        return

    xls = pd.ExcelFile(excel_path)
    df_datos = pd.read_excel(xls, sheet_name="Datos")
    df_resumen = pd.read_excel(xls, sheet_name="Resumen")
    df_conceptos = pd.read_excel(xls, sheet_name="Conceptos")

    nuevas_filas = []
    
    if saldo_ant is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ANTERIOR", "MONTO": saldo_ant})
    
    filas_originales = df_resumen.to_dict('records')
    filas_limpias = [f for f in filas_originales if f["CONCEPTO"] not in ["SALDO ANTERIOR", "SALDO ACTUAL"]]
    
    nuevas_filas.extend(filas_limpias)

    if saldo_nue is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ACTUAL", "MONTO": saldo_nue})

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


def ejecutar_proceso_exportacion(pdf_path, output_excel_path=None, progreso_callback=None):
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

    lineas_plana, saldo_ant, saldo_nue = extraer_lineas_davivienda(pdf_path, progreso_callback)

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