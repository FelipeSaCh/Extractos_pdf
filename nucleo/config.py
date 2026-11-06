import os
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent

SERVIDOR = r"\\SRV-CONTAB\Servidor de Archivos\2025 SERVIDOR\APLICATIVO FELIPE"
RUTA_ADMINISTRADORAS = SERVIDOR + r"\Documentos_auditor\administradoras.json"
RUTA_ACTUALIZACIONES = SERVIDOR + r"\Extractor PDF"

CARPETA_USUARIO = Path(os.getenv("APPDATA") or Path.home()) / "ExtractorExcel"
RUTA_CONFIG_PLANO = CARPETA_USUARIO / "config_plano.json"

URL_MANUAL = "https://drive.google.com/file/d/1ZB5-3KT_-2bJmo7G-uYIlwJhnN-BihQP/view?usp=sharing"


def ruta_asset(*partes):
    return str(RAIZ.joinpath("assets", *partes))