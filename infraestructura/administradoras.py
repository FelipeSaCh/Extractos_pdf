import json
import os
import shutil
import tempfile

import pandas as pd

from nucleo.config import RUTA_ADMINISTRADORAS
from nucleo.texto import codigo, digitos, normalizar


def leer(ruta=RUTA_ADMINISTRADORAS):
    """(json original, clave de la lista o None, registros, mtime)."""
    with open(ruta, encoding="utf-8-sig") as f:
        raw = json.load(f)
    clave = None
    if isinstance(raw, dict):
        listas = [k for k, v in raw.items() if isinstance(v, list) and (not v or isinstance(v[0], dict))]
        clave = next((k for k in listas if raw[k]), next(iter(listas), None))
        if clave is None:
            raise ValueError("administradoras.json no contiene una lista de administradoras.")
        registros = raw[clave]
    elif isinstance(raw, list):
        registros = raw
    else:
        raise ValueError("Formato de administradoras.json no reconocido.")
    return raw, clave, [dict(r) for r in registros if isinstance(r, dict)], os.path.getmtime(ruta)


def modificado_desde(ruta, mtime):
    return os.path.getmtime(ruta) != mtime


def guardar(ruta, raw, clave, registros):
    """Respaldo .bak y escritura atómica; devuelve el nuevo mtime."""
    contenido = registros
    if clave is not None:
        raw[clave] = registros
        contenido = raw
    shutil.copy2(ruta, ruta + ".bak")
    fd, temporal = tempfile.mkstemp(suffix=".tmp", dir=os.path.dirname(ruta))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(contenido, f, ensure_ascii=False, indent=2)
        os.replace(temporal, ruta)
    except Exception:
        if os.path.exists(temporal):
            os.remove(temporal)
        raise
    return os.path.getmtime(ruta)


def tabla(ruta=RUTA_ADMINISTRADORAS):
    """Administradoras normalizadas: nombre, nit, categoria, cuenta, codigo."""
    df = pd.DataFrame(leer(ruta)[2])
    cols = {normalizar(c): c for c in df.columns}

    def col(clave, obligatoria=True):
        c = next((o for n, o in cols.items() if clave in n), None)
        if c is None:
            if obligatoria:
                raise ValueError(f"administradoras.json no contiene la columna '{clave}'.")
            return pd.Series("", index=df.index)
        return df[c]

    return pd.DataFrame({
        "nombre": col("NOMBRE").map(normalizar),
        "nit": col("NIT").map(digitos),
        "categoria": col("CATEGOR", False).map(normalizar),
        "cuenta": col("CUENTA").map(digitos),
        "codigo": col("CODIGO", False).map(lambda v: "" if pd.isna(v) else codigo(v)),
    })