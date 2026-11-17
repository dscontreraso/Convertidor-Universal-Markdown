# Convertidor Universal a Markdown (GUI)

Aplicación de escritorio en Python (`Tkinter`) que escanea de forma recursiva cualquier directorio, todas sus subcarpetas y archivos `.zip` (incluyendo `.zip` anidados dentro de otros `.zip`) para convertir todos los documentos encontrados en archivos Markdown (`.md`) estructurados.

## Características Principales

- **Escaneo Recursivo Profundo**: Recorre todas las subcarpetas físicas y explora el interior de archivos `.zip` y `.zip` anidados sin necesidad de descomprimirlos manualmente.
- **Contadores en Tiempo Real**: Muestra el total de archivos (directos y en ZIP), total de subcarpetas (físicas y dentro de ZIP), cantidad de archivos `.zip` y desglose por formato.
- **Explorador Jerárquico Integrado**: Árbol interactivo para inspeccionar las subcarpetas y archivos detectados antes de iniciar la conversión.
- **Procesamiento Multihilo**: Conversión paralela configurable (`workers`) con barra de progreso, velocidad en `archivos/s`, tiempo estimado restante y generación automática de `INFORME_CONVERSION.md`.

## Formatos Soportados

| Categoría | Extensiones | Extracción a Markdown |
| :--- | :--- | :--- |
| **PDF** | `.pdf` | Texto por página, metadatos y detección/formateo de tablas (`PyMuPDF`) |
| **Excel / Datos** | `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv` | Conversión de todas las hojas en tablas Markdown (`pandas` + `tabulate`) |
| **Word** | `.docx`, `.doc`, `.rtf`, `.odt` | Jerarquía de títulos, párrafos, listas y tablas (`python-docx`) |
| **PowerPoint** | `.pptx`, `.ppt` | Texto por diapositiva, tablas y notas del orador (`python-pptx`) |
| **Imágenes** | `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tif` | Dimensiones, modo de color, DPI, tamaño y enlace Markdown (`Pillow`) |
| **Texto / Código** | `.txt`, `.json`, `.xml`, `.html`, `.py`, `.sql`, `.log`, etc. | Bloques de código con resaltado de sintaxis y formateo JSON |
| **Archivos `.ZIP`** | `.zip` (incluye `.zip` dentro de `.zip`) | Inspección y conversión recursiva en memoria |

## Instalación y Uso

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/TU_USUARIO/Convertidor-Universal-Markdown.git
   cd Convertidor-Universal-Markdown
   ```

2. Instalar las dependencias:
   ```bash
   pip install -r requirements.txt
   ```

3. Ejecutar la aplicación:
   ```bash
   python convertidor_markdown_app.py
   ```

---

## Autor

Desarrollado por **Daniel Contreras**  
Contacto: [dscontreraso@unal.edu.co](mailto:dscontreraso@unal.edu.co)

