"""Registro único de fuentes (bancos y operadores de planilla) que consume la interfaz."""
import os
from dataclasses import dataclass, field

from servicios.libro_extracto import escribir_libro, exportar

BTN_EXTRACTO = "Extracto"
BTN_PN = "MovimientosPNatural"
BTN_SOC = "MovimientoSOC"
BTN_PLANILLA = "SeguridadSocial"
BTN_EXCEL = "excel"
BTN_PLANO = "plano"

ETIQUETAS_EXISTENTES = (
    "extracto_", "movimiento_", "movimientos_", "movimientospnatural_", "movimientosoc_", "seguridadsocial_",
)
SIN_REGISTROS = "No se obtuvieron registros del documento procesado."


def _bancolombia():
    from extractores import bancolombia
    return bancolombia


def _bogota():
    from extractores import bogota
    return bogota


def _nequi():
    from extractores import nequi
    return nequi


def _davivienda():
    from extractores import davivienda
    return davivienda


def _arus():
    from servicios import seguridad_social
    return seguridad_social


def _aportes_en_linea():
    from servicios import aportes_en_linea
    return aportes_en_linea


@dataclass(frozen=True)
class Fuente:
    nombre: str
    cargador: object
    botones: tuple
    planilla: bool = False
    acepta_excel: bool = False
    _cache: dict = field(default_factory=dict, compare=False, repr=False)

    def modulo(self):
        """Importa el módulo de la fuente; el error queda asociado a esta fuente."""
        if "modulo" not in self._cache:
            try:
                self._cache["modulo"] = self.cargador()
            except Exception as e:
                raise ImportError(f"No se pudo cargar el módulo de {self.nombre}:\n{e}") from e
        return self._cache["modulo"]

    def opciones(self):
        """{parametro: (titulo, texto, {clave: descripcion}, detectar(ruta))} que el usuario debe elegir."""
        return getattr(self.modulo(), "OPCIONES", {})

    def generador(self):
        return self.modulo().Generador

    def exportar(self, ruta, destino=None, progreso=None, **opciones):
        """Extrae el documento y genera el libro de Excel; devuelve su ruta."""
        modulo = self.modulo()
        extraccion = modulo.extraer(ruta, progreso, **opciones)
        if not extraccion.lineas:
            raise ValueError(getattr(modulo, "SIN_MOVIMIENTOS", SIN_REGISTROS))
        return exportar(extraccion, ruta, destino)


FUENTES = {f.nombre: f for f in (
    Fuente("Bancolombia (Estándar)", _bancolombia, (BTN_EXTRACTO, BTN_PN, BTN_SOC, BTN_EXCEL, BTN_PLANO)),
    Fuente("Banco de Bogotá (Extractos PyME)", _bogota, (BTN_EXTRACTO,)),
    Fuente("Nequi", _nequi, (BTN_EXTRACTO,)),
    Fuente("Davivienda", _davivienda, (BTN_EXTRACTO,)),
    Fuente("Seguridad social (ARUS)", _arus, (BTN_PLANILLA,), planilla=True),
    Fuente("Seguridad social (Aportes en Línea)", _aportes_en_linea, (BTN_PLANILLA,), planilla=True, acepta_excel=True),
)}
SIN_FUENTE = "-- Selecciona un banco --"


def con_prefijo(ruta, prefijo):
    """Renombra el archivo con el prefijo del tipo de documento si aún no tiene uno reconocido."""
    carpeta, nombre = os.path.split(ruta)
    if nombre.lower().startswith(ETIQUETAS_EXISTENTES):
        return ruta
    nueva = os.path.join(carpeta, f"{prefijo}_{nombre}")
    os.replace(ruta, nueva)
    return nueva


def procesar_excel_bancolombia(entrada, salida=None):
    from extractores import bancolombia_excel
    base, ext = os.path.splitext(entrada)
    salida = salida or f"{base}_procesado{ext if ext.lower() == '.xlsx' else '.xlsx'}"
    e = bancolombia_excel.extraer(entrada)
    return escribir_libro(e.dataframe(), salida, e.saldo_anterior, e.saldo_actual)