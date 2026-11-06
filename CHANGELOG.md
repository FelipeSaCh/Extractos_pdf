# Changelog

Todos los cambios notables en este proyecto serán documentados en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/)
y este proyecto se adhiere a [Semantic Versioning](https://semver.org/lang/es/).

## [2.7.0] - 2026-11-06
### Añadido
- Al cargar una planilla de seguridad social (ARUS o Aportes en Línea) se pregunta si es de **empleados** o de **independiente**.
- Planilla de independiente: salud → 519605001, pensión → 519605002, ARL → 519605003. Intereses de mora, CCF y la contrapartida conservan sus cuentas.
- Identificación del subsistema por código PILA (EPS…, 23xxxx/25-14, 14-xx, CCF…) o, en su defecto, por la categoría de administradoras.json.
- 7 pruebas nuevas (107 en total).

## [2.6.0] - 2026-11-06
### Modificado
- Arquitectura por capas: `nucleo`, `infraestructura`, `extractores`, `servicios` y `presentacion`. `app_gui.py` queda como punto de entrada.
- Registro único de fuentes (`servicios/fuentes.py`): reemplaza las importaciones con try/except y las cadenas if/elif de la interfaz; cada fuente reporta su propio error de carga.
- Un solo convertidor de montos (`nucleo/numeros.a_float`) reemplaza siete funciones; un solo intérprete de fechas (`nucleo/fechas.partes_fecha`).
- Lectura de PDF en una sola apertura por documento (`leer_paginas`), incluida la detección de la cuenta bancaria; progreso por página en todos los bancos.
- El libro de Excel se escribe en una sola pasada desde memoria (antes: escribir, releer y reescribir).
- «Cargar Excel» usa el mismo generador de libro que los PDF (resúmenes por mes, conciliación y columna `DCTO.`).
- Persistencia de administradoras y de parámetros del plano separada de la interfaz.
- La conversión a Excel 5.0/95 importa pywin32 sólo al generar el plano.

### Añadido
- Módulo de pruebas automatizadas (`tests/`, 100 pruebas con pytest; `pip install -r requirements-dev.txt` y `python -m pytest`).

### Corregido
- Banco de Bogotá: la marca «FIN MOVIMIENTOS» no se detectaba y el pie de página se pegaba al último movimiento.
- Movimientos Persona Natural: montos con punto de miles sin decimales (`-$50.000`) se leían como -50.
- Fechas «dd mmm aaaa» (Persona Natural) quedaban sin año/mes/día en el archivo plano.
- El tipo de documento Bancolombia se toma del botón usado, no sólo del nombre del archivo.
- `process_pdf` podía fallar con variable sin definir; mensaje de error de importación incorrecto entre bancos.
- `minsize` mayor que la ventana inicial; `requirements.txt` sin numpy ni pywin32.

### Eliminado
- `clasificador_conceptos.py`, `movimientos_test.py` y `app_gui.spec` movidos a `_legacy/` (sin uso).

## [2.1.0] - 2026-08-19
### Añadido
- Soporte para procesamiento de extractos de nequi. 
### Modificado
- Actualización de la versión global a `2.1.0` en `version.py`.

## [2.0.0] - 2026-08-19
### Añadido
- Soporte para procesar archivos PDF protegidos con contraseña, solicitándola dinámicamente mediante una ventana de diálogo en la GUI.
- Procesamiento en segundo plano utilizando `worker_manager.py` para evitar que la interfaz gráfica se congele (no responsiva) durante la carga de PDFs grandes y la extracción de datos.
- Módulo `ui_theme.py` que centraliza los estilos visuales, colores de la interfaz y componentes comunes como el cuadro de advertencias (Warning Box).
- Barra de progreso que muestra el avance en tiempo real tanto del renderizado del PDF como de la exportación de movimientos.
- Nuevo botón y funcionalidad de "Limpiar sesión" para liberar la memoria y desbloquear archivos abiertos en el sistema operativo.

### Modificado
- Refactorización de la lógica del motor en `pdf_engine.py` para separar renderizado de páginas e integrar callbacks de progreso de forma compatible con hilos secundarios.
- Actualización de la versión global a `2.0.0` en `version.py`.

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