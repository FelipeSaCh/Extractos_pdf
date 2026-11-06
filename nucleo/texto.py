import math
import re
import unicodedata


def sin_tildes(texto):
    return unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().upper()


def normalizar(texto):
    return re.sub(r"[^A-Z0-9]+", " ", sin_tildes(texto)).strip()


def clave(texto):
    return normalizar(texto).replace(" ", "")


def codigo(texto):
    return re.sub(r"\s", "", str(texto or "")).upper()


def digitos(valor):
    if isinstance(valor, float):
        valor = "" if math.isnan(valor) else int(valor)
    return re.sub(r"\D", "", str(valor if valor is not None else ""))