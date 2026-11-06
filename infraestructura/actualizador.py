"""Actualizaciones de Extractor PDF desde el servidor (sin dependencias de interfaz)."""
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from nucleo.config import RUTA_ACTUALIZACIONES
from version import __version__

APP = "ExtractorPDF"
PREFIJO_INSTALADOR = "ExtractosPDF_Setup"
ARCHIVO_VERSION = "version.json"
ARCHIVO_RUTA_LOCAL = "actualizacion.json"
ARGS_SILENCIOSO = ("/SILENT", "/SUPPRESSMSGBOXES", "/NORESTART", "/CLOSEAPPLICATIONS")


def tupla(v: str) -> tuple:
    return tuple(int(p) if p.isdigit() else 0 for p in str(v).strip().split("."))


def es_compilado() -> bool:
    import __main__
    return bool(getattr(sys, "frozen", False) or getattr(sys, "_MEIPASS", None)
                or "__compiled__" in globals() or hasattr(__main__, "__compiled__"))


def sha256(ruta) -> str:
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def _ruta_local() -> str | None:
    try:
        with open(Path(sys.argv[0]).resolve().parent / ARCHIVO_RUTA_LOCAL, encoding="utf-8-sig") as f:
            return json.load(f).get("ruta") or None
    except (OSError, ValueError, AttributeError):
        return None


def ruta_servidor() -> Path:
    return Path(_ruta_local() or RUTA_ACTUALIZACIONES)


def leer_publicada(carpeta: Path) -> dict | None:
    try:
        with open(carpeta / ARCHIVO_VERSION, encoding="utf-8-sig") as f:
            return json.load(f)
    except FileNotFoundError:
        return None


def buscar() -> dict | None:
    carpeta = ruta_servidor()
    info = leer_publicada(carpeta)
    if not info or tupla(info.get("version", "0")) <= tupla(__version__):
        return None
    instalador = carpeta / info["instalador"]
    if not instalador.exists():
        raise FileNotFoundError(f"No se encontró el instalador publicado: {instalador}")
    return {**info, "ruta_instalador": str(instalador)}


def instalar(info: dict) -> None:
    carpeta = Path(os.getenv("LOCALAPPDATA") or tempfile.gettempdir()) / APP / "actualizaciones"
    carpeta.mkdir(parents=True, exist_ok=True)
    for viejo in carpeta.glob("*.exe"):
        try:
            viejo.unlink()
        except OSError:
            pass
    destino = carpeta / Path(info["ruta_instalador"]).name
    shutil.copy2(info["ruta_instalador"], destino)
    if info.get("sha256") and sha256(destino) != info["sha256"]:
        destino.unlink(missing_ok=True)
        raise IOError("El instalador copiado está incompleto o dañado. Verifica la conexión a la red e intenta de nuevo.")
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
    subprocess.Popen([str(destino), *ARGS_SILENCIOSO], creationflags=flags, close_fds=True, cwd=str(carpeta))