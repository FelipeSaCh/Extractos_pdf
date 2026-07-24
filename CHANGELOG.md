# Changelog

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/)
y este proyecto se adhiere a [Semantic Versioning](https://semver.org/lang/es/).

## [1.0.0] - 2026-07-24
### Añadido
- Control de versiones estructurado del software (`version.py`).
- Mención de la versión (`v1.0.0`) en la barra de título de la aplicación GUI.
- Documentación e indicaciones sobre la preparación para el empaquetado a ejecutable `.exe` (el programa está listo para ser empaquetado, aunque el archivo ejecutable compilado aún no está distribuido/disponible directamente).

### Modificado
- Normalización de nombres de columnas en la exportación de Excel (`DESCRIPCION` sin tilde) para evitar problemas de compatibilidad y facilitar consultas.
- Lógica de reorganización de Excel mejorada. Ahora distribuye los datos procesados en tres hojas dinámicas dentro del mismo archivo para un análisis contable más limpio:
  - `Datos`: Transacciones ordenadas dejando los cargos (valores negativos) en la parte superior.
  - `Resumen`: Totales de cargos (egresos) y abonos (ingresos).
  - `Conceptos`: Agrupación y totalización de transacciones por descripción.
