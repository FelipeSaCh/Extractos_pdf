"""Parámetros del archivo plano recordados por NIT (%APPDATA%\\ExtractorExcel\\config_plano.json)."""
import json
import os

from nucleo.config import RUTA_CONFIG_PLANO


def cargar():
    try:
        with open(RUTA_CONFIG_PLANO, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def guardar(cfg):
    try:
        os.makedirs(os.path.dirname(RUTA_CONFIG_PLANO), exist_ok=True)
        with open(RUTA_CONFIG_PLANO, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


def guardar_parametros(resultado):
    """Guarda comprobantes por clase y el próximo consecutivo de cada comprobante para el NIT."""
    cfg = cargar()
    empresa = cfg.setdefault("empresas", {}).setdefault(resultado["nit"], {})
    empresa.setdefault("clases", {}).update(resultado["clases"])
    consecutivos = empresa.setdefault("consecutivos", {})
    for tipo, inicio in resultado["consecutivos"].items():
        consecutivos[tipo] = inicio + int(resultado["conteo"].get(tipo, 0))
    cfg["ultimo_nit"] = resultado["nit"]
    guardar(cfg)