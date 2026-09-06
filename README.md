# 🌐 Sistema de Traducción por Lotes con Agentes LangChain + Gemini

<div align="center">

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://translation-batch-app-ctletvzn6rlgmqv8cngwrl.streamlit.app/)
[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![Streamlit](https://img.shields.io/badge/Streamlit-1.38+-FF4B4B?style=for-the-badge&logo=streamlit&logoColor=white)](https://streamlit.io)
[![LangChain](https://img.shields.io/badge/LangChain-LCEL-1C3C3C?style=for-the-badge&logo=langchain&logoColor=white)](https://www.langchain.com/)
[![Google Gemini](https://img.shields.io/badge/Google%20Gemini-Flash--Lite-4285F4?style=for-the-badge&logo=google&logoColor=white)](https://aistudio.google.com/)
[![Docker](https://img.shields.io/badge/Docker-Enabled-2496ED?style=for-the-badge&logo=docker&logoColor=white)](https://www.docker.com/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

**Plataforma web de traducción inteligente de documentos por lotes (`.txt`, `.docx`, `.pdf`) con pipeline de agentes autónomos, verificación cruzada visual por colores y exportación masiva en ZIP.**

👉 **[Probar la aplicación en vivo (Live Demo)](https://translation-batch-app-ctletvzn6rlgmqv8cngwrl.streamlit.app/)**

</div>

---

## 📌 Tabla de Contenidos

1. [Características Principales](#-características-principales)
2. [Arquitectura de los 4 Agentes](#-arquitectura-de-los-4-agentes)
3. [Despliegue en Streamlit Community Cloud](#-despliegue-en-streamlit-community-cloud)
4. [Ejecución con Docker](#-ejecución-con-docker-recomendado-en-local)
5. [Ejecución Local con Python](#-ejecución-local-sin-docker)
6. [Flujo de Uso y Funcionalidades](#-flujo-de-uso-y-funcionalidades)
7. [Variables de Entorno y Configuración](#-variables-de-entorno-y-configuración)
8. [Estructura del Repositorio](#-estructura-del-repositorio)
9. [Resiliencia y Manejo de Errores](#-resiliencia-y-manejo-de-errores)
10. [Licencia](#-licencia)

---

## ✨ Características Principales

- 📑 **Traducción por lotes multiformato**: Soporte integral para archivos `.txt`, `.docx` y `.pdf`.
- 🤖 **Pipeline de 4 agentes autónomos**: Extractor, Traductor (Gemini Flash-Lite), Validador Heurístico y Alineador Semántico.
- 🎨 **Interfaz visual premium (Glassmorphism)**: Diseño dark mode moderno con tipografía *Outfit* y *Plus Jakarta Sans*, tarjetas de métricas interactivas y contenedores con scroll interno independiente.
- 🔍 **Verificación cruzada sincronizada por colores**: Mapeo visual 1 a 1 entre segmentos originales y traducidos para una auditoría rápida y sin esfuerzo.
- ✏️ **Editor de segmentos integrado**: Posibilidad de corregir manualmente cualquier segmento antes de exportar.
- 📦 **Exportación individual y masiva**: Descarga individual por archivo en su formato original/elegido o descarga global del lote en un único archivo `.zip`.
- 🛡️ **Rate-limiting y Backoff Exponencial**: Control proactivo de llamadas por minuto y reintentos inteligentes ante cuotas 429 de la API de Google Gemini.
- 🧠 **Auditoría y Memoria de Pipeline**: Registro completo de eventos y decisiones de los agentes mediante `ConversationBufferMemory`.

---

## 🏛️ Arquitectura de los 4 Agentes

```
[ Archivo(s) de Entrada (.txt, .docx, .pdf) ]
                      │
                      ▼
┌──────────────────────────────────────────────┐
│  1. AGENTE EXTRACTOR                         │
│  - Parseo por formato                        │
│  - Limpieza y segmentación inteligente       │
└──────────────────────┬───────────────────────┘
                      │  Lista de Segmentos (original)
                      ▼
┌──────────────────────────────────────────────┐
│  2. AGENTE TRADUCTOR (LangChain + Gemini)    │
│  - Traducción determinista LCEL              │
│  - RateLimiter & Backoff exponencial         │
└──────────────────────┬───────────────────────┘
                      │  Segmentos traducidos
                      ▼
┌──────────────────────────────────────────────┐
│  3. AGENTE VALIDADOR                         │
│  - Análisis de longitud y retención léxica   │
│  - Reintentos automáticos si es sospechoso   │
└──────────────────────┬───────────────────────┘
                      │  Segmentos validados
                      ▼
┌──────────────────────────────────────────────┐
│  4. AGENTE ALINEADOR                         │
│  - Asignación de paleta cromática cíclica    │
│  - Validación semántica (Embeddings opcional)│
└──────────────────────┬───────────────────────┘
                      │
                      ▼
[ Vista Previa Interactiva + Exportación .TXT/.DOCX/.PDF / .ZIP ]
```

---

## 🚀 Despliegue en Streamlit Community Cloud

Para desplegar la aplicación en la nube de Streamlit de forma gratuita:

1. Haz un **Fork** o sube este repositorio a tu cuenta de GitHub.
2. Inicia sesión en [share.streamlit.io](https://share.streamlit.io/).
3. Haz clic en **"New app"** y selecciona:
   - **Repository:** `tu-usuario/translation_batch_app`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. En **"Advanced settings" ➔ "Secrets"**, añade tu API Key de Gemini:
   ```toml
   GEMINI_API_KEY = "tu_api_key_de_gemini_aqui"
   GEMINI_MODEL = "gemini-3.1-flash-lite"
   DEFAULT_SOURCE_LANG = "auto"
   DEFAULT_TARGET_LANG = "en"
   ```
5. Haz clic en **"Deploy"**. ¡La aplicación estará en línea en segundos!

---

## 🐳 Ejecución con Docker (Recomendado en Local)

### 1. Clonar el repositorio y configurar variables
```bash
git clone https://github.com/Prolexis/translation-batch-app.git
cd translation_batch_app
cp .env.example .env
```
Edita el archivo `.env` y coloca tu API Key de Gemini:
```ini
GEMINI_API_KEY=tu_api_key_aqui
```

### 2. Levantar el contenedor con Docker Compose
```bash
docker compose up --build
```
Abre tu navegador en: **`http://localhost:8501`**

Para detener el contenedor:
```bash
docker compose down
```

---

## 💻 Ejecución Local (Sin Docker)

Si prefieres ejecutar directamente en tu entorno Python local:

```bash
# 1. Crear entorno virtual
python -m venv venv

# 2. Activar entorno
# En Windows (PowerShell):
venv\Scripts\Activate.ps1
# En Linux/macOS:
source venv/bin/activate

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Configurar variables de entorno
cp .env.example .env

# 5. Ejecutar la app
streamlit run app.py
```

---

## 📖 Flujo de Uso y Funcionalidades

1. **Configuración en Barra Lateral**:
   - **API Key**: Detectada automáticamente desde `.env` o ingresada manualmente de forma segura.
   - **Idioma Origen**: `🌐 Detección automática (auto)` o selección manual de idioma.
   - **Idioma Destino**: Idioma al cual se traducirá el lote (por defecto `🇬🇧 Inglés`).
   - **Modo de Alineación**: `position` (rápido por párrafo) o `semantic` (análisis con Embeddings).
   - **Formato de Exportación**: `.docx`, `.pdf` o `.txt`.
2. **Carga de Documentos**: Arrastra uno o múltiples archivos.
3. **Procesar Lote**: Haz clic en **🚀 Procesar Lote** para traducir en tiempo real.
4. **Resultados y Verificación**:
   - Inspecciona las métricas de cada archivo (Segmentos, Tasa de advertencia, Tiempo).
   - Compara en el visor de doble columna con scroll sincronizado por color.
   - Edita segmentos puntuales si lo deseas en el editor colapsable.
5. **Descarga**: Exporta individualmente cada documento o pulsa **📦 Descargar Lote Completo (.ZIP)**.

---

## ⚙️ Variables de Entorno y Configuración

| Variable | Valor por Defecto | Descripción |
|---|---|---|
| `GEMINI_API_KEY` | *(Requerido)* | Clave de API de Google Gemini |
| `GEMINI_MODEL` | `gemini-3.1-flash-lite` | Modelo de lenguaje utilizado para traducción |
| `GEMINI_EMBEDDING_MODEL` | `models/text-embedding-004` | Modelo para validación semántica |
| `DEFAULT_SOURCE_LANG` | `auto` | Idioma de origen por defecto |
| `DEFAULT_TARGET_LANG` | `en` | Idioma objetivo por defecto |
| `MAX_CALLS_PER_MINUTE` | `30` | Límite preventivo de llamadas salientes |
| `BACKOFF_BASE_SECONDS` | `2.0` | Tiempo base para reintentos con backoff |
| `BACKOFF_MAX_RETRIES` | `5` | Número máximo de reintentos ante error 429 |
| `LENGTH_DIFF_THRESHOLD` | `0.40` | Umbral de discrepancia de longitud (±40%) |
| `MAX_RETRIES` | `2` | Reintentos por segmento del Agente Validador |

---

## 📁 Estructura del Repositorio

```
translation_batch_app/
├── app.py                     # Aplicación Streamlit completa y orquestador
├── config.py                  # Dataclass de configuración y variables de entorno
├── agents/
│   ├── extractor_agent.py     # Agente Extractor de texto
│   ├── translator_agent.py    # Agente Traductor (LCEL + Gemini)
│   ├── validator_agent.py     # Agente Validador de calidad
│   ├── aligner_agent.py       # Agente Alineador (posición / embeddings)
│   └── orchestrator.py        # Orquestador del pipeline
├── utils/
│   ├── file_handlers.py       # Handlers para .txt, .docx y .pdf
│   ├── colors.py              # Paleta cromática y renderizador HTML
│   └── rate_limiter.py        # RateLimiter y decorador with_backoff
├── sample_files/              # Documentos de prueba (.txt, .docx, .pdf)
├── Dockerfile                 # Imagen Docker optimizada
├── docker-compose.yml         # Despliegue con un comando
├── requirements.txt           # Dependencias de Python
├── .env.example               # Plantilla de variables de entorno
├── LICENSE                    # Licencia MIT
└── README.md                  # Documentación del proyecto
```

---

## 🛡️ Resiliencia y Manejo de Errores

- **Límites de Cuota (429 / Quota Exceeded)**: Decorador `@with_backoff` con jitter aleatorio para absorber picos de tasa de llamadas.
- **Archivos Dañados o PDFs sin OCR**: Captura granular con `CorruptFileError` y `UnsupportedFormatError` sin detener el procesamiento del resto del lote.
- **Validación Heurística de Traducción**: Detección automática de textos no traducidos o discrepancias anómalas con reintento automático.

---

## 📄 Licencia

Este proyecto está bajo la Licencia **MIT**. Consulta el archivo [LICENSE](file:///c:/Users/Usuario/Documents/ghitub/translation_batch_app/LICENSE) para más detalles.
