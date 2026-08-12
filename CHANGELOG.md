# Changelog

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/)
y este proyecto se adhiere a [Semantic Versioning](https://semver.org/lang/es/).

## [1.5.8] - 2026-08-12
### Añadido
- Nuevo módulo `script_ban_bogota.py` diseñado para la extracción y procesamiento de extractos en PDF de **Banco de Bogotá (Extractos PyME)**.
- Selector desplegable ("Selecciona tu banco") en la interfaz gráfica (`app_gui.py`) para dar soporte a múltiples bancos de forma estructurada.
- Habilitación dinámica de botones y lógica en la GUI según el banco seleccionado (actualmente limitando a carga de extractos para Banco de Bogotá).
- Integración de los flujos de procesamiento y exportación de archivos tanto para Bancolombia como para Banco de Bogotá en la aplicación gráfica.

### Modificado
- Actualizada la advertencia de bancos soportados en la interfaz de usuario para incluir al Banco de Bogotá.
- Actualizada la versión global del programa a `1.5.8` en `version.py`.

## [1.1.8] - 2026-08-11
### Añadido

- Nuevo modulo de movimientos
- Ventana de multi hojas para lectura de excel procesados despues de su conversion
- Se pone "Movimimientos en mantenimiento para correccion de bugs"
- Se añade saldo anterior y saldo actual a la hoja de resumen
- Se añade "Limpiar sesion" para liberar memoria y archivos ocupados por la app
- Se añade settings.json
- Boton de carga de excel
- Boton para cargar movimiento de persona natural
- Transformacion para movimiento de persona natural
- Se adiciona funcionalidad para trasnformacion de extractos originales en excel
- Nueva ventana emergente para especificar informacion bancaria y organizacion de Dataframe
- Columna para conceptos por "Mes"

### Modificado
-Se modifican formatos
-Se añade un totalizado para los abonos y cargos
- Cambios visuales para app_gui.py
- Script de back para manejar la lectura de movimientos
- Se corrige recalculacion erronea al re-organizar Dataframe
- Se elimina boton de "Reorganizar" y se integra funcion al crear .xlsx/.tx
- Cambio de nombre para boton de movimiento a Movimiento Sociedades
- Fix: movimientos de P.Natural con texto fuera de la informacion contable
- En la hoja conceptos, se dejan lineas separadas por cada mes