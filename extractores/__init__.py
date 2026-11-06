"""Extractores por banco. Cada módulo expone extraer(ruta, progreso=None, **opciones) -> Extraccion."""
from dataclasses import dataclass

import pandas as pd

from nucleo.constantes import DELIMITADOR


@dataclass
class Extraccion:
    columnas: list
    lineas: list
    cuenta: str = ""
    saldo_anterior: float | None = None
    saldo_actual: float | None = None

    def texto(self):
        return "\n".join([DELIMITADOR.join(self.columnas), *self.lineas])

    def dataframe(self):
        partes = (linea.split(DELIMITADOR) for linea in self.lineas)
        return pd.DataFrame([p for p in partes if len(p) == len(self.columnas)], columns=self.columnas)