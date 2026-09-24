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
    """Procesa el PDF reconociendo la tabla y extrayendo saldos del recuadro inicial."""
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
        
        # ELIMINADO: Ya no forzamos year = "2024". Si no lo encuentra, quedará como None.

        # Buscar "Saldo Anterior" y "Nuevo Saldo" palabra por palabra para mayor precisión
        words_pag1 = primer_pag.extract_words()
        for i, w in enumerate(words_pag1):
            txt = w['text'].upper()
            
            # Detectar Saldo Anterior
            if txt == "SALDO" and i + 1 < len(words_pag1) and words_pag1[i+1]['text'].upper() == "ANTERIOR":
                for j in range(i+2, min(i+6, len(words_pag1))):
                    val_txt = words_pag1[j]['text'].replace('$', '').replace(',', '')
                    if re.match(r'^-?[\d\.]+$', val_txt):
                        saldo_anterior = float(val_txt)
                        break
                        
            # Detectar Nuevo Saldo
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

                # --- NUEVO: Exclusiones ampliadas de pie de página ---
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
                
                # Si encuentra alguna de estas frases del aviso legal, corta la lectura de esa página
                if any(frase in texto_linea for frase in frases_pie_pagina):
                    break 
                
                if "FECHA DÍA MES" in texto_linea or "OFICINA" in texto_linea or "DESCRIPCIÓN" in texto_linea or "CRÉDITO" in texto_linea:
                    continue
                # -------------------------------------------------------
                
                es_nueva_tx = False
                if len(palabras_linea) >= 2:
                    dia_text = palabras_linea[0]['text'].strip()
                    mes_text = palabras_linea[1]['text'].strip()
                    if re.match(r'^\d{1,2}$', dia_text) and re.match(r'^\d{1,2}$', mes_text) and palabras_linea[0]['x0'] < anc * 0.10:
                        es_nueva_tx = True

                if es_nueva_tx:
                    if tx_actual:
                        _guardar_tx(tx_actual, lineas_delimitadas)

                    # --- NUEVO: Lógica dinámica de la fecha ---
                    if year:
                        fecha_formateada = f"{dia_text.zfill(2)}/{mes_text.zfill(2)}/{year}"
                    else:
                        fecha_formateada = f"{dia_text.zfill(2)}/{mes_text.zfill(2)}"

                    tx_actual = {
                        "FECHA": fecha_formateada,
                        "OFICINA": [],
                        "DESCRIPCIÓN": [],
                        "DOCUMENTO": [],
                        "DEBITO": [],
                        "CREDITO": []
                    }

                    for w in palabras_linea[2:]:
                        x_mid = (w['x0'] + w['x1']) / 2.0
                        if x_mid < anc * 0.16:
                            tx_actual["OFICINA"].append(w['text'])
                        elif x_mid < anc * 0.58:
                            tx_actual["DESCRIPCIÓN"].append(w['text'])
                        elif x_mid < anc * 0.68:
                            tx_actual["DOCUMENTO"].append(w['text'])
                        elif x_mid < anc * 0.83:
                            tx_actual["DEBITO"].append(w['text'])
                        else:
                            tx_actual["CREDITO"].append(w['text'])
                else:
                    if tx_actual and palabras_linea:
                        x_primera = palabras_linea[0]['x0']
                        if anc * 0.16 <= x_primera < anc * 0.68:
                            for w in palabras_linea:
                                x_mid = (w['x0'] + w['x1']) / 2.0
                                if x_mid < anc * 0.58:
                                    tx_actual["DESCRIPCIÓN"].append(w['text'])
                                elif x_mid < anc * 0.68:
                                    tx_actual["DOCUMENTO"].append(w['text'])

            if tx_actual:
                _guardar_tx(tx_actual, lineas_delimitadas)
                tx_actual = None

            if progreso_callback:
                progreso_callback(indice_pagina, total_paginas)

    return lineas_delimitadas, saldo_anterior, nuevo_saldo

def _guardar_tx(tx_dict, lista_salida):
    """Calcula un único VALOR financiero unificando Débito y Crédito."""
    desc = " ".join(tx_dict["DESCRIPCIÓN"]).strip()
    oficina = " ".join(tx_dict["OFICINA"]).strip()
    doc = " ".join(tx_dict["DOCUMENTO"]).strip()

    debito_str = " ".join(tx_dict["DEBITO"]).replace('$', '').replace(',', '').strip()
    credito_str = " ".join(tx_dict["CREDITO"]).replace('$', '').replace(',', '').strip()

    try:
        deb_val = float(debito_str) if debito_str else 0.0
    except ValueError:
        deb_val = 0.0

    try:
        cred_val = float(credito_str) if credito_str else 0.0
    except ValueError:
        cred_val = 0.0

    if deb_val > 0:
        valor_final = f"-{deb_val}"
    elif cred_val > 0:
        valor_final = str(cred_val)
    else:
        valor_final = "0.0"

    registro = [
        tx_dict["FECHA"],
        "",
        desc,
        "",
        oficina,
        doc,
        valor_final,
        ""
    ]
    lista_salida.append(DELIMITADOR.join(registro))

def inyectar_saldos_en_resumen(excel_path, saldo_ant, saldo_nue):
    """Inserta los saldos encontrados directamente en la hoja Resumen, limpiando duplicados falsos."""
    if saldo_ant is None and saldo_nue is None:
        return

    # Leemos las hojas existentes
    xls = pd.ExcelFile(excel_path)
    df_datos = pd.read_excel(xls, sheet_name="Datos")
    df_resumen = pd.read_excel(xls, sheet_name="Resumen")
    df_conceptos = pd.read_excel(xls, sheet_name="Conceptos")

    nuevas_filas = []
    
    # 1. Colocar el Saldo Anterior real arriba del todo (si existe)
    if saldo_ant is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ANTERIOR", "MONTO": saldo_ant})
    
    # 2. Filtrar las filas basura que calculó script.py (quitamos sus versiones de Saldo)
    filas_originales = df_resumen.to_dict('records')
    filas_limpias = [f for f in filas_originales if f["CONCEPTO"] not in ["SALDO ANTERIOR", "SALDO ACTUAL"]]
    
    # Agregar solo TOTAL CARGOS y TOTAL ABONOS
    nuevas_filas.extend(filas_limpias)

    # 3. Colocar el Nuevo Saldo real al final (si existe)
    if saldo_nue is not None:
        nuevas_filas.append({"CONCEPTO": "SALDO ACTUAL", "MONTO": saldo_nue})

    df_resumen_actualizado = pd.DataFrame(nuevas_filas)

    # Sobrescribir el Excel regenerando el formato visual
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

    # Desempaquetamos los tres valores que ahora retorna la función
    lineas_plana, saldo_ant, saldo_nue = extraer_lineas_davivienda(pdf_path, progreso_callback)

    if not lineas_plana:
        raise ValueError("No se encontraron movimientos del extracto o el formato es irreconocible.")

    if not output_excel_path:
        base_path, _ = os.path.splitext(pdf_path)
        output_excel_path = f"{base_path}_convertido.xlsx"

    # 1. Guardar archivo Plano (.txt)
    base_txt_path, _ = os.path.splitext(output_excel_path)
    txt_path = f"{base_txt_path}_plano.txt"

    encabezado_txt = DELIMITADOR.join(columnas)
    contenido_txt = [encabezado_txt] + lineas_plana

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("\n".join(contenido_txt))

    # 2. Convertir a Excel
    texto_delimitado_a_excel(lineas_plana, columnas, output_excel_path)
    
    # 3. Aplicar la reestructuración estética de pestañas base (Genera Datos, Resumen y Conceptos)
    reorganizar_excel(output_excel_path)
    
    # 4. Inyectar los saldos obtenidos de la portada directamente en la hoja "Resumen"
    inyectar_saldos_en_resumen(output_excel_path, saldo_ant, saldo_nue)

    return output_excel_path