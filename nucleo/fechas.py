import datetime
import math
import re

import pandas as pd

from nucleo.constantes import MESES

_ABREV_MES = {m[:3].lower(): i for i, m in enumerate(MESES, start=1)}
_ISO = re.compile(r"^(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})")
_CORTA = re.compile(r"^(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2}|\d{4}))?(?!\d)")
_TEXTO = re.compile(rf"^(\d{{1,2}})\s+({'|'.join(_ABREV_MES)})[a-z]*\.?(?:\s+(\d{{4}}))?", re.I)
_NULA = (math.nan, math.nan, math.nan)


def partes_fecha(valor):
    """(año, mes, día) de una fecha del extracto: datetime, 'aaaa/mm/dd', 'dd/mm[/aa]', 'dd mmm [aaaa]' o d.mm numérico."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return _NULA
    if isinstance(valor, (datetime.date, pd.Timestamp)):
        return valor.year, valor.month, valor.day
    texto = f"{valor:.2f}" if isinstance(valor, float) else str(valor).strip()
    if m := _ISO.match(texto):
        return int(m[1]), int(m[2]), int(m[3])
    if m := _CORTA.match(texto):
        ano = int(m[3]) + 2000 if m[3] and int(m[3]) < 100 else int(m[3]) if m[3] else math.nan
        return ano, int(m[2]), int(m[1])
    if m := _TEXTO.match(texto):
        return int(m[3]) if m[3] else math.nan, _ABREV_MES[m[2][:3].lower()], int(m[1])
    return _NULA


def mes_de_fechas(fechas):
    mes = pd.to_numeric(fechas.map(lambda v: partes_fecha(v)[1]), errors="coerce")
    return mes.where(mes.between(1, 12)).fillna(0).astype(int)