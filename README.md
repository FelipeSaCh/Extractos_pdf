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

La solución está desarrollada bajo una estructura modular en **Python 3**:

* **`app_gui.py`**: Interfaz gráfica creada con `tkinter`. Controla la selección de banco, la interacción del usuario y hospeda los componentes visuales. Incorpora hilos secundarios (`worker_manager.py`) para evitar el bloqueo del hilo principal al procesar y renderizar PDFs.
* **`pdf_engine.py`**: Motor encargado de renderizar y manipular las páginas del PDF visualmente con `PyMuPDF` (`fitz`), incluyendo soporte para descifrar archivos protegidos por contraseña.
* **`script.py`**: Procesador de extractos de **Bancolombia** mediante lectura geométrica con `pdfplumber` y segmentación espacial de columnas.
* **`script_ban_bogota.py`**: Procesador específico para extractos de **Banco de Bogotá (PyME)**, empleando anclaje dinámico por fechas y coordenadas optimizadas.
* **`clasificador_conceptos.py`**: Módulo interactivo que permite clasificar contablemente los movimientos directamente desde la GUI del Excel generado.
* **`parser_excel.py`**: Lector de archivos de Excel procesados para su integración y visualización en el panel del programa.
* **`ui_theme.py`**: Centraliza el estilo visual de la interfaz gráfica (colores, fuentes y elementos comunes como el cuadro de advertencias).
* **`worker_manager.py`**: Administrador de hilos secundarios en segundo plano para Tkinter, evitando que la aplicación quede "No responde".
* **`requirements.txt`**: Lista de dependencias del proyecto.

---

## 🚀 Requisitos e Instalación

### Requisitos previos
* Python 3.8 o superior instalado en el sistema.

### Instalación de Dependencias
1. Clona este repositorio o descarga los archivos.
2. Abre una terminal en la raíz del proyecto.
3. Instala las librerías necesarias con el siguiente comando:
   ```bash
   pip install -r requirements.txt
   ```
   *Nota: Adicionalmente, para el renderizado del visor del PDF, se requieren las librerías `PyMuPDF` (`fitz`) y `Pillow` (PIL), las cuales deben estar instaladas en el entorno.*

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

* **Versión Actual:** `v2.0.0`
* **Historial de Cambios:** Todos los cambios detallados de cada versión se encuentran en el archivo [CHANGELOG.md](file:///c:/Users/USUARIO/Desktop/Proyectos/extractos_pdf/CHANGELOG.md).

### 📦 Compilación a Ejecutable (.exe)
El código de la versión actual está estructurado y optimizado para ser empaquetado en un archivo ejecutable `.exe` independiente (utilizando herramientas como `PyInstaller`). 

> **Estado Actual:** El empaquetado y la distribución directa del binario `.exe` **no están disponibles todavía** para descarga en este repositorio (actualmente en desarrollo). 
> 
> Si deseas generar el archivo ejecutable de forma local, puedes preparar el entorno e iniciar la compilación con el siguiente comando:
> ```bash
> pip install pyinstaller
> pyinstaller --noconfirm --onefile --windowed --name="LectorExtractorBancolombia" app_gui.py
> ```
