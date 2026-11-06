"""Libro de Excel del extracto: hojas Datos, Resumen, Conceptos e Info."""
import os

import pandas as pd

from infraestructura.excel import aplicar_formato_moneda, anexar_resumenes, hoja_info, negrita
from nucleo.constantes import HOJA_DATOS, HOJA_INFO, MESES
from nucleo.fechas import mes_de_fechas
from nucleo.numeros import serie_a_numero


def _resumen(valores, saldos=None, saldo_anterior=None, saldo_actual=None):
    conciliar = saldo_anterior is not None and saldo_actual is not None
    if saldos is not None and not saldos.empty:
        saldo_anterior = saldos.iloc[0] - valores.iloc[0] if saldo_anterior is None else saldo_anterior
        saldo_actual = saldos.iloc[-1] if saldo_actual is None else saldo_actual
    cargos, abonos = valores[valores < 0].sum(), valores[valores > 0].sum()
    filas = [("SALDO ANTERIOR", saldo_anterior)] if saldo_anterior is not None else []
    filas += [("TOTAL CARGOS", cargos), ("TOTAL ABONOS", abonos)]
    if saldo_actual is not None:
        filas.append(("SALDO ACTUAL", saldo_actual))
    if conciliar:
        filas.append(("DIFERENCIA CONCILIACIÓN", round(saldo_anterior + cargos + abonos - saldo_actual, 2)))
    return pd.DataFrame(filas, columns=["CONCEPTO", "MONTO"])


def _bloque_resumen(columnas, col_desc, titulo, resumen):
    return pd.DataFrame({
        col_desc: [titulo, *resumen["CONCEPTO"]],
        "VALOR": [None, *resumen["MONTO"]],
    }).reindex(columns=columnas)


def _intercalar_resumenes(columnas, col_desc, bloques, resumen_general):
    partes, negritas, fila = [], [], 2

    def agregar(df, en_negrita=False):
        nonlocal fila
        partes.append(df)
        if en_negrita:
            negritas.extend(range(fila, fila + len(df)))
        fila += len(df)

    for titulo, resumen, filas in bloques:
        agregar(filas)
        agregar(pd.DataFrame(index=range(1), columns=columnas))
        agregar(_bloque_resumen(columnas, col_desc, titulo, resumen), True)
        agregar(pd.DataFrame(index=range(2), columns=columnas))
    agregar(_bloque_resumen(columnas, col_desc, "RESUMEN GENERAL", resumen_general), True)
    return pd.concat(partes, ignore_index=True), negritas


def _conceptos(meses, descripciones, valores, col_desc):
    calc = pd.DataFrame({
        "MES": meses,
        col_desc: descripciones,
        "CARGOS": valores.where(valores < 0, 0.0),
        "ABONOS": valores.where(valores > 0, 0.0),
        "NETO": valores,
    })
    agrupado = calc.groupby(["MES", col_desc], as_index=False).sum().sort_values(["MES", "CARGOS"])
    vacia = dict.fromkeys(calc.columns)
    filas = []
    for i, (_, grupo) in enumerate(agrupado.groupby("MES", sort=True)):
        filas += [dict(vacia) for _ in range(3)] if i else []
        filas += grupo.to_dict("records")
    filas += [dict(vacia) for _ in range(3)]
    filas.append({**vacia, col_desc: "TOTAL GENERAL", "CARGOS": float(agrupado["CARGOS"].sum()),
                  "ABONOS": float(agrupado["ABONOS"].sum()), "NETO": float(agrupado["NETO"].sum())})
    return pd.DataFrame(filas, columns=list(vacia))


def escribir_libro(df, excel_path, saldo_anterior=None, saldo_actual=None, cuenta_bancaria=""):
    """Escribe en una sola pasada el libro organizado a partir de los movimientos (texto) en memoria."""
    df = df.astype(object).mask(df.astype(str).apply(lambda s: s.str.strip()).eq("")).reset_index(drop=True)
    col_desc = "DESCRIPCIÓN" if "DESCRIPCIÓN" in df.columns else "DESCRIPCION"
    valores = serie_a_numero(df["VALOR"]).fillna(0.0)
    tiene_saldo = "SALDO" in df.columns and df["SALDO"].notna().any()
    saldos = serie_a_numero(df["SALDO"]).fillna(0.0) if tiene_saldo else None
    meses = mes_de_fechas(df["FECHA"]) if "FECHA" in df.columns else pd.Series(0, index=df.index)
    resumen = _resumen(valores, saldos, saldo_anterior, saldo_actual)
    conceptos = _conceptos(meses, df[col_desc].fillna(""), valores, col_desc)

    datos = df.assign(VALOR=valores)
    if tiene_saldo:
        datos["SALDO"] = saldos
    orden_meses = list(pd.unique(meses))
    negritas = []
    if len(orden_meses) > 1:
        bloques = [
            (f"RESUMEN {MESES[mes - 1]}" if mes else "RESUMEN SIN FECHA",
             _resumen(valores[meses == mes], saldos[meses == mes] if tiene_saldo else None),
             datos[meses == mes])
            for mes in orden_meses
        ]
        datos, negritas = _intercalar_resumenes(datos.columns, col_desc, bloques, resumen)
    else:
        datos = datos.loc[(valores >= 0).sort_values(kind="stable").index]

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        datos.to_excel(writer, sheet_name=HOJA_DATOS, index=False)
        resumen.to_excel(writer, sheet_name="Resumen", index=False)
        conceptos.to_excel(writer, sheet_name="Conceptos", index=False)
        if cuenta_bancaria:
            hoja_info(cuenta_bancaria).to_excel(writer, sheet_name=HOJA_INFO, index=False)
        libro = writer.book
        for ws in libro.worksheets:
            aplicar_formato_moneda(ws)
        col_concepto, col_monto = datos.columns.get_loc(col_desc) + 1, datos.columns.get_loc("VALOR") + 1
        if negritas:
            negrita(libro[HOJA_DATOS], negritas, (col_concepto, col_monto))
        else:
            anexar_resumenes(libro[HOJA_DATOS], [("RESUMEN", resumen)], col_concepto, col_monto)
        anexar_resumenes(libro["Conceptos"], [("RESUMEN", resumen)],
                         conceptos.columns.get_loc(col_desc) + 1, conceptos.columns.get_loc("NETO") + 1)
    return excel_path


def exportar(extraccion, ruta_origen, output_excel_path=None, respaldo_txt=True):
    """Guarda el respaldo _plano.txt y el libro del extracto; devuelve la ruta del .xlsx."""
    if not extraccion.lineas:
        return None
    output_excel_path = output_excel_path or f"{os.path.splitext(ruta_origen)[0]}_convertido.xlsx"
    if respaldo_txt:
        with open(f"{os.path.splitext(output_excel_path)[0]}_plano.txt", "w", encoding="utf-8") as f:
            f.write(extraccion.texto())
    return escribir_libro(extraccion.dataframe(), output_excel_path,
                          extraccion.saldo_anterior, extraccion.saldo_actual, extraccion.cuenta)