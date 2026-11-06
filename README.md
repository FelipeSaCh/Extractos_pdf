# Extractor de excel para informacion bancaria

Este proyecto nace con la finalidad de automatizar y optimizar la extracción de transacciones desde extractos bancarios en formato PDF, convirtiéndolos en formatos estructurados como Excel (`.xlsx`) y texto plano delimitado (`.txt`). Actualmente, la herramienta es compatible con extractos de **Bancolombia** y **Banco de Bogotá (Extractos PyME)**.

---

## 📌 Planteamiento de la Necesidad

> **"Extraer información necesaria de extractos bancarios (en su primera versión, customizado para extractos bancolombia). Al trabajar con extractos de clientes, existen casos en donde se tiene que volver a solicitar el envío del mismo, al tener poco acceso de comunicación o integración con el cliente, por lo que para evitar dichos re-procesos, se hace el planteamiento inicial de dicha herramienta."**

### Ampliación del Contexto y Necesidad
En procesos contables, de auditoría, conciliación bancaria y análisis financiero, el insumo principal son los extractos bancarios del cliente. Sin embargo, en el día a día se presentan los siguientes desafíos:
1. **Dificultades de comunicación:** A menudo, los canales con los clientes son lentos o indirectos, lo que hace que cualquier solicitud de información tome días o semanas.
2. **Formatos inconsistentes o ilegibles:** Los clientes pueden enviar archivos incompletos, capturas de pantalla o reportes dañados, lo que tradicionalmente obliga a solicitar reenvíos.
3. **Falta de integraciones directas:** Para pequeñas y medianas empresas o clientes independientes, no existen APIs bancarias de fácil acceso que permitan descargar la información de forma automática.
4. **Re-procesos y digitación manual:** Cuando la información es difícil de procesar, se incurre en errores de digitación y en la pérdida de tiempo valioso al transcribir manualmente cada movimiento de filas interminables de PDFs de decenas de páginas.

Para solucionar estos problemas y evitar el re-proceso de solicitar constantemente nuevos documentos o procesarlos manualmente, esta herramienta ofrece una forma rápida de **extraer de forma local, segura y estructurada** toda la información del extracto en segundos.

---

## ⚙️ Funcionamiento de la Aplicación

La herramienta consta de una interfaz gráfica de escritorio (GUI) interactiva que guía al usuario a través del siguiente flujo de trabajo:

1. **Selección del Banco:** El usuario selecciona el banco emisor del extracto mediante el menú desplegable ("Bancolombia (Estándar)" o "Banco de Bogotá (Extractos PyME)"). Los controles de la aplicación se habilitarán dinámicamente de acuerdo con el soporte del banco seleccionado.
2. **Carga del PDF:** El usuario selecciona el archivo del extracto en formato PDF. La aplicación procesa internamente el documento y muestra una **vista previa visual** en tiempo real de las páginas en el panel derecho.
3. **Extracción y Conversión:**
   - La aplicación analiza espacialmente el contenido de cada página (coordenadas de texto) utilizando algoritmos de reconocimiento de columnas y posiciones relativas específicas de cada banco.
   - Genera un archivo **plano de texto** (`*_plano.txt`) con los campos delimitados por el carácter especial `|||`.
   - Genera una **hoja de cálculo de Excel (`.xlsx`)** estructurada con las columnas correspondientes del banco procesado (e.g. `FECHA`, `DESCRIPCIÓN`, `SUCURSAL`, `DCTO.`, `VALOR` y `SALDO` para Bancolombia, o `FECHA`, `COD TRANS`, `DESCRIPCIÓN`, `CIUDAD`, `OFICINA/CANAL`, `DOCUMENTO`, `VALOR`, `SALDO` para Banco de Bogotá).
4. **Reorganización Inteligente e Informe:** Adicionalmente, el sistema genera de forma automática pestañas suplementarias de **Resumen** (conciliación de saldo anterior, saldo actual, abonos y cargos) y **Conceptos** (agrupaciones automáticas de movimientos por mes y descripción con totales generales).
5. **Visualización Directa:** Con la interfaz visual, el usuario puede lanzar y clasificar conceptos directamente o utilizar *Abrir Excel* para ver el archivo final en el software de hojas de cálculo predeterminado (por ejemplo, Microsoft Excel).

---

## 🛠️ Arquitectura de Desarrollo

Arquitectura por capas en **Python 3.10+**. Cada capa sólo depende de las inferiores:

```
presentacion  →  servicios  →  extractores / infraestructura  →  nucleo
```

| Capa | Contenido |
|---|---|
| `nucleo/` | Constantes, rutas (`config.py`), conversión de montos y fechas, normalización de texto, validaciones. Sin E/S. |
| `infraestructura/` | PDF (`pdf_engine.py`, `lectura_pdf.py`), Excel y conversión 5.0/95 (`excel.py`), `administradoras.json`, actualizador. |
| `extractores/` | Un módulo por banco con `extraer(ruta, progreso, **opciones) -> Extraccion`. |
| `servicios/` | Registro de fuentes, libro del extracto, archivo plano, seguridad social (ARUS / Aportes en Línea), parámetros por NIT. |
| `presentacion/` | Interfaz Tkinter: ventana principal, parámetros del plano, editor de administradoras, hilos y diálogos. |

`app_gui.py` es el punto de entrada. Para agregar un banco: crear `extractores/<banco>.py` con `extraer()` y registrarlo en `servicios/fuentes.py`.

---

## 🚀 Requisitos e Instalación

### Requisitos previos
* Python 3.10 o superior instalado en el sistema.

### Instalación de Dependencias
1. Clona este repositorio o descarga los archivos.
2. Abre una terminal en la raíz del proyecto.
3. Instala las librerías necesarias con el siguiente comando:
   ```bash
   pip install -r requirements.txt
   ```

---

## 💻 Instrucciones de Uso

Para iniciar la aplicación, ejecuta el siguiente comando en la consola:

```bash
python app_gui.py
```

Una vez abierta la interfaz:
1. Haz clic en **📁 Cargar PDF** y selecciona tu extracto Bancolombia.
2. Visualiza el extracto en el visualizador.
3. Haz clic en **⚡ Extraer TXT y Excel** para generar los archivos estructurados en la ruta que desees.
4. Si necesitas analizar egresos primero, haz clic en **🔄 Reorganizar Excel** (esto moverá los egresos a la parte superior).
5. Haz clic en **🟢 Abrir Excel** para revisar el resultado final.

---

## 📌 Control de Versiones y Empaquetado

Este proyecto adopta **Semantic Versioning (SemVer)** para gestionar las versiones del programa de forma ordenada y siguiendo las mejores prácticas de la industria.

* **Versión Actual:** `v2.7.0`
* **Historial de Cambios:** Todos los cambios detallados de cada versión se encuentran en el archivo [CHANGELOG.md](file:///c:/Users/USUARIO/Desktop/Proyectos/extractos_pdf/CHANGELOG.md).

### 📦 Compilación y publicación
1. Subir la versión en `version.py`.
2. `compilar.bat`: compila con Nuitka (standalone) y genera el instalador con Inno Setup en `instaladores\ExtractosPDF_Setup_<versión>.exe`.
3. `publicar.bat`: copia el instalador al servidor con verificación SHA-256 y actualiza `version.json`.