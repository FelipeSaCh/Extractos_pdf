# Extractor de excel para informacion bancaria

Este proyecto nace con la finalidad de automatizar y optimizar la extracción de transacciones desde extractos bancarios en formato PDF, convirtiéndolos en formatos estructurados como Excel (`.xlsx`) y texto plano delimitado (`.txt`). En esta primera versión, la herramienta está personalizada y optimizada específicamente para los extractos de **Bancolombia**.

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

1. **Carga del PDF:** El usuario selecciona el archivo del extracto en formato PDF. La aplicación procesa internamente el documento y muestra una **vista previa visual** en tiempo real de las páginas en el panel derecho.
2. **Extracción y Conversión:**
   - La aplicación analiza espacialmente el contenido de cada página (coordenadas de texto) utilizando algoritmos de reconocimiento de columnas.
   - Genera un archivo **plano de texto** (`*_plano.txt`) con los campos delimitados por el carácter especial `|||`.
   - Genera una **hoja de cálculo de Excel** (`*_convertido.xlsx`) con las columnas estructuradas: `FECHA`, `DESCRIPCIÓN`, `SUCURSAL`, `DCTO.`, `VALOR` y `SALDO`.
3. **Reorganización Inteligente (Opcional):** Permite, mediante un botón, reordenar el Excel de forma que todos los movimientos con montos negativos (gastos/débitos) se posicionen en la parte superior para facilitar auditorías de flujo de caja y revisión de egresos de forma rápida.
4. **Visualización Directa:** Con el botón *Abrir Excel*, el usuario puede lanzar el archivo generado directamente en el software de hojas de cálculo predeterminado del sistema operativo (por ejemplo, Microsoft Excel o LibreOffice).

---

## 🛠️ Arquitectura de Desarrollo

La solución está desarrollada bajo una estructura modular en **Python 3**:

* **`app_gui.py`**: Interfaz gráfica creada con `tkinter` (el framework estándar de GUI para Python). Maneja los eventos de usuario, diálogos de selección de archivos y la representación visual de las páginas del PDF.
* **`pdf_engine.py`**: Motor encargado de la manipulación visual del PDF. Utiliza `PyMuPDF` (`fitz`) para renderizar las páginas a objetos de imagen compatibles con Tkinter.
* **`script.py`**: Núcleo del procesamiento y lógica de negocio. Utiliza `pdfplumber` para leer la posición geométrica exacta de cada palabra en el PDF y agruparlas en filas y columnas basadas en umbrales de posición (`x0` y `x1`), reduciendo al mínimo los errores de salto de línea típicos de extractos con diseños complejos.
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

* **Versión Actual:** `v1.0.0`
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
