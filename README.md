# Convertidor Universal a Markdown (GUI + OCR + LaTeX)

Aplicación de escritorio en Python (`Tkinter`) que escanea de forma recursiva cualquier directorio, todas sus subcarpetas y archivos `.zip` (incluyendo `.zip` anidados dentro de otros `.zip`) para convertir todos los documentos encontrados en archivos Markdown (`.md`) estructurados.

## Características Principales

- **Respeto de Formato en PDF**:
  - **Encabezados Jerárquicos (`##`, `###`, `####`)**: Analiza tamaños de fuente, pesos tipográficos (negrita) y numeración estructural (`CAPÍTULO`, `ARTÍCULO`, `1.`, `1.1.`, `1.1.1.`).
  - **Tablas en Formato Limpio**: Extrae tablas en su posición vertical exacta de lectura, limpia saltos de línea internos, escapa caracteres `|`, elimina filas/columnas vacías y alinea las columnas automáticamente.
  - **Ecuaciones en LaTeX (`$$...$$` y `$...$`)**: Detecta ecuaciones independientes (con soporte de numeración `\tag{n}`), superíndices, subíndices, raíces, fracciones y símbolos matemáticos/griegos, además de ecuaciones nativas OMML (`m:oMath`) en documentos Word (`.docx`).
- **OCR Integrado (`RapidOCR`) para Documentos Escaneados e Imágenes**:
  - Detecta automáticamente páginas de PDF escaneadas (sin texto digital) y aplica OCR de alta velocidad con reconstrucción de encabezados, tablas multicolumna y ecuaciones en LaTeX.
  - Extrae texto, tablas y fórmulas directamente desde archivos de imagen (`.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tif`).
- **Escaneo Recursivo Profundo**: Recorre todas las subcarpetas físicas y explora el interior de archivos `.zip` y `.zip` anidados sin necesidad de descomprimirlos manualmente.
- **Contadores en Tiempo Real**: Muestra el total de archivos (directos y en ZIP), total de subcarpetas (físicas y dentro de ZIP), cantidad de archivos `.zip` y desglose por formato.
- **Procesamiento Multihilo**: Conversión paralela configurable (`workers`) con barra de progreso, velocidad en `archivos/s`, tiempo estimado restante y generación automática de `INFORME_CONVERSION.md`.

## Formatos Soportados

| Categoría | Extensiones | Extracción a Markdown |
| :--- | :--- | :--- |
| **PDF (Digital y Escaneado)** | `.pdf` | Encabezados tipográficos, tablas limpias en orden de lectura, ecuaciones LaTeX y **OCR automático** (`PyMuPDF` + `RapidOCR`) |
| **Word** | `.docx`, `.doc`, `.rtf`, `.odt` | Jerarquía de títulos, párrafos, listas, tablas limpias y ecuaciones `OMML` a LaTeX (`python-docx`) |
| **Excel / Datos** | `.xlsx`, `.xls`, `.xlsm`, `.csv`, `.tsv` | Conversión de todas las hojas en tablas Markdown limpias y alineadas (`pandas`) |
| **PowerPoint** | `.pptx`, `.ppt` | Texto por diapositiva, ecuaciones en LaTeX, tablas limpias y notas del orador (`python-pptx`) |
| **Imágenes (con OCR)** | `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tif` | Metadatos + extracción completa de texto, tablas y ecuaciones vía **RapidOCR** (`Pillow` + `RapidOCR`) |
| **Texto / Código** | `.txt`, `.json`, `.xml`, `.html`, `.py`, `.sql`, `.log`, `.tex`, etc. | Bloques de código con resaltado de sintaxis y formateo JSON |
| **Archivos `.ZIP`** | `.zip` (incluye `.zip` dentro de `.zip`) | Inspección y conversión recursiva en memoria |

## Instalación y Uso

1. Clonar el repositorio:
   ```bash
   git clone https://github.com/dscontreraso/Convertidor-Universal-Markdown.git
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
