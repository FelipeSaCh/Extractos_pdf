import math
import re

import pandas as pd

_DECIMAL_FINAL = re.compile(r"[.,](\d{1,2})$")
_SIGNO_NEGATIVO = re.compile(r"^[^\d]*-|-\s*$|^\s*\(.*\)\s*$")


def a_float(valor, defecto=0.0):
    """Convierte montos en formato latino o anglosajón ("1.234.567,89", "-$1,234.56", "50.000", "(1.200)").

    Un separador seguido de 1-2 dígitos al final es decimal; los demás son de miles.
    """
    if valor is None or isinstance(valor, bool):
        return defecto if valor is None else float(valor)
    if isinstance(valor, (int, float)):
        return defecto if isinstance(valor, float) and math.isnan(valor) else float(valor)
    texto = str(valor).strip()
    cifras = re.sub(r"[^\d.,]", "", texto)
    m = _DECIMAL_FINAL.search(cifras)
    entero, decimal = (cifras[:m.start()], m[1]) if m else (cifras, "")
    entero = re.sub(r"\D", "", entero)
    if not entero and not decimal:
        return defecto
    numero = float(f"{entero or 0}.{decimal or 0}")
    return -numero if _SIGNO_NEGATIVO.search(texto) else numero


def serie_a_numero(serie):
    """Versión vectorial de a_float; deja NaN en lo no convertible."""
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    return serie.map(lambda v: a_float(v, math.nan)).astype(float)


def a_entero(valor, minimo=1):
    try:
        return max(int(str(valor).strip()), minimo)
    except (TypeError, ValueError):
        return minimo