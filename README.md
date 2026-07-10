# Project Report Assistance

Automatiza la creación de informes técnicos de certificación (`.docx`) a partir de los datos de un ensayo: extrae la información de las distintas fuentes, reconoce el marcado de los componentes mediante la **API de Claude (Anthropic)** y arma el informe final según la norma correspondiente.

## ¿Qué hace?

A partir de una carpeta de trabajo con las fotos correctamente renombradas, archivos del sistema, de la certificadora y un .xlsx con indicaciones básica, el programa:

1. **Extrae** los datos del sistema, la certificadora y el archivo .xlsx.
2. **Reconoce el marcado** de cada componente en las fotos usando visión por IA (OCR con la API de Claude).
3. **Genera el informe** `.docx` seleccionando automáticamente la plantilla según la norma, y completa marcados, cláusulas, tablas e imágenes.

## Uso de la API (clave del proyecto)

El reconocimiento del marcado de componentes se hace con la **API de Anthropic (Claude)** en modo visión, en [`extract_05_api_ocr_marcado.py`](Main/01%20Extractors/extract_05_api_ocr_marcado.py). Puntos importantes:

- **Modelo**: `claude-sonnet-4-5` (configurable en la constante `MODEL`).
- **Optimización de costos**: las imágenes se redimensionan y recomprimen antes de enviarse. El script **estima el costo** antes de llamar a la API y **aborta si supera el umbral** `COST_THRESHOLD` (por defecto `0.03` USD).
- **Prompts especializados**: hay un prompt por tipo de componente (capacitores, fusibles, varistores, etc.) para leer con precisión símbolos como `µ`, `~`, `Ω`, `²`, `±` y evitar confusiones (0/O, 1/I, µ/u).
- **Control de tasa**: pausa configurable entre llamadas (`DELAY_BETWEEN_CALLS`) y reintento automático ante `RateLimitError`.
- **Salida**: guarda lo reconocido en `output/json/extracted_api_ocr_marking.json`.

### Configurar la clave de API

El script busca la clave en este orden:

1. Variable de entorno `ANTHROPIC_API_KEY` (recomendado):

   ```powershell
   setx ANTHROPIC_API_KEY "sk-ant-tu-clave-aqui"
   ```

   (Reiniciá la terminal después de configurarla.)

2. Como alternativa, un archivo local con la clave.

> ⚠️ **Nunca subas tu clave de API al repositorio.** Mantené la clave fuera del control de versiones (usá la variable de entorno o un archivo ignorado por Git).

## Requisitos

- Python 3.10+
- Dependencias: `anthropic`, `opencv-python-headless`, `Pillow`, `numpy`, `pandas`, `openpyxl`, `python-docx` (el script de OCR instala las suyas automáticamente).

## Estructura del proyecto

```
Main/
  00/                 Orquestadores y utilidades (main_path, all_stages, ...)
  01 Extractors/      Extracción de datos + OCR por API
  02 Doc modifiers/   Generación y modificación del .docx por norma
.XLSX .DOCX patrones/ Plantillas de informe y lógica de cláusulas
```

## Cómo ejecutar

1. Indicá la carpeta de trabajo en [`Main/00/main_path.py`](Main/00/main_path.py):

   ```python
   main_path = r"C:\ruta\a\tu\carpeta_de_trabajo"
   ```

   La carpeta debe contener las fotos (con la nomenclatura correspondiente), `sistema.htm`, `certificadora.pdf` y el `.xlsx` de EUT/componentes configurado para dicho componente.

2. Ejecutá el pipeline completo:

   ```powershell
   python "Main/00/all_stages.py"
   ```

   Esto corre la extracción (`extract_00_all.py`) y luego selecciona y ejecuta el generador de informe según la norma (`standard_selection.py`).

El resultado (informe `.docx`, fotos renombradas y JSON intermedios) queda en la subcarpeta `output/` dentro de la carpeta de trabajo.


