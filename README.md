# 🌐 Sistema de Traducción por Lotes con Agentes LangChain + Gemini

Aplicación web (Streamlit) que traduce documentos `.txt`, `.docx` y `.pdf` en
lote, con revisión en vivo, edición manual y verificación cruzada por
colores, orquestada mediante un pipeline de 4 agentes construidos sobre
LangChain y la API de Gemini.

---

## 1. Requisitos previos

- Docker y Docker Compose (recomendado), **o** Python 3.11+ si se ejecuta local.
- Una API Key de Google Gemini: https://aistudio.google.com/app/apikey

## 2. Ejecución rápida con Docker (recomendado)

```bash
# 1. Clona/descomprime el proyecto y entra a la carpeta
cd translation_batch_app

# 2. Copia el archivo de entorno de ejemplo y coloca tu API Key
cp .env.example .env
nano .env        # reemplaza GEMINI_API_KEY=coloca_aqui_tu_api_key

# 3. Construye y levanta el contenedor
docker compose up --build

# 4. Abre la aplicación
# http://localhost:8501
```

Para detener: `docker compose down`.
La API Key también puede pegarse directamente en el campo de la barra
lateral de la app si prefieres no usar el archivo `.env`.

### Ejecución con `docker run` (sin compose)

```bash
docker build -t batch-translation-agents .
docker run --rm -p 8501:8501 --env-file .env batch-translation-agents
```

## 3. Ejecución local sin Docker (opcional)

```bash
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env              # completa tu GEMINI_API_KEY
streamlit run app.py
```

## 4. Uso de la aplicación

1. En la barra lateral, confirma/edita tu `GEMINI_API_KEY`, idioma origen,
   idioma destino, modo de alineación (`position` o `semantic`) y formato
   de exportación.
2. Sube 2 o más archivos (`.txt`, `.docx`, `.pdf`) en el cargador de
   archivos. Aparecerán en la **tabla de estado** como "Pendiente".
3. Pulsa **"🚀 Procesar lote"**. Verás la barra de progreso global y,
   apenas cada archivo termina, su resultado se despliega de inmediato
   (revisión en vivo) sin esperar a que termine el lote completo.
4. Para cada archivo, se muestran dos columnas: **original** y
   **traducido**, ambas resaltadas con colores; el segmento N del original
   y el segmento N de la traducción comparten exactamente el mismo color.
5. El texto traducido es editable directamente en los cuadros de texto
   antes de exportar.
6. Usa el botón **"⬇️ Descargar traducción"** para exportar cada archivo en
   el formato elegido (`.txt`, `.docx` o `.pdf`).
7. El expander **"🧠 Historial de agentes"** muestra la traza completa de lo
   que hizo cada agente (útil para auditoría/depuración).

Se incluyen 3 archivos de prueba en `sample_files/` (uno por formato) para
probar la demo de inmediato.

## 5. Estructura del proyecto

```
translation_batch_app/
├── app.py                     # Interfaz Streamlit (UI + orquestación de alto nivel)
├── config.py                  # Configuración centralizada (env vars, umbrales, colores)
├── agents/
│   ├── extractor_agent.py     # Agente Extractor
│   ├── translator_agent.py    # Agente Traductor (LangChain + Gemini)
│   ├── validator_agent.py     # Agente Validador (criterios medibles + reintentos)
│   ├── aligner_agent.py       # Agente de Alineación (posición o embeddings)
│   └── orchestrator.py        # Orquestador (equivalente a AgentExecutor/LCEL)
├── utils/
│   ├── file_handlers.py       # Extracción/exportación .txt/.docx/.pdf
│   ├── colors.py               # Paleta cíclica + render HTML resaltado
│   └── rate_limiter.py         # Rate limiting + backoff exponencial (429)
├── sample_files/               # 3 archivos de prueba (.txt, .docx, .pdf)
├── requirements.txt
├── Dockerfile
├── docker-compose.yml
├── .env.example
└── ARCHITECTURE.md              # Documento de arquitectura de agentes (entregable)
```

Ver **`ARCHITECTURE.md`** para el detalle de roles, entradas/salidas y
justificación de cada agente (entregable de documentación pedido en el
enunciado).

## 6. Manejo de errores contemplado

| Escenario | Manejo |
|---|---|
| Formato de archivo no soportado | `UnsupportedFormatError`, se marca el archivo como "error" en la tabla sin detener el resto del lote. |
| Archivo corrupto / PDF cifrado / docx dañado | `CorruptFileError`, capturada en el Agente Extractor. |
| PDF escaneado sin texto extraíble | Se informa el error explícitamente ("posible PDF sin OCR"). |
| Error 429 / cuota excedida de Gemini | `utils/rate_limiter.with_backoff`: reintenta con backoff exponencial + jitter (hasta `BACKOFF_MAX_RETRIES`). |
| Timeout / servicio no disponible (503) | Mismo mecanismo de backoff que 429. |
| Traducción sospechosa (longitud o vocabulario) | Agente Validador solicita reintento al Traductor (máx. `MAX_RETRIES` por segmento). |
| Fallo irrecuperable en un segmento | Se marca `status="error"`, se contabiliza en la tasa de error del archivo, no bloquea los demás segmentos/archivos. |
| API Key ausente/ inválida | La UI bloquea el botón "Procesar lote" y muestra un mensaje explícito. |

## 7. Notas y límites conocidos

- El extractor de PDF depende de `pypdf`; PDFs generados solo con líneas de
  texto sin marcas de párrafo (p. ej. un PDF armado a mano con `canvas` de
  bajo nivel) pueden extraerse como un único bloque. PDFs exportados desde
  Word/Google Docs/LibreOffice conservan la estructura de párrafos
  correctamente.
- El modo de alineación `semantic` consume cuota adicional de la API
  (embeddings) y es más lento; se recomienda para lotes pequeños o cuando
  se sospeche reordenamiento de segmentos.
- La app está pensada para uso de un único usuario por contenedor (el
  estado vive en `st.session_state`, propio de la sesión del navegador).
