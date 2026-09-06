# 📄 Arquitectura del Sistema de Marcado y Trazabilidad de Origen para Papers Académicos

Este documento técnico explica la arquitectura, el modelo de datos y el ciclo de vida de los identificadores de posición y metadatos de trazabilidad en la aplicación de traducción académica con agentes LangChain y Google Gemini (`gemini-3.1-flash-lite`).

---

## 1. El Problema de la Trazabilidad en Traducción Científica

Cuando un investigador traduce un artículo científico o paper académico del inglés al español para utilizarlo en su tesis, estado del arte o publicación, enfrenta un desafío crítico: **la pérdida de procedencia**. 
Si el traductor simplemente devuelve un bloque continuo de texto en español:
- Es imposible saber de qué página o sección del documento original proviene un párrafo específico.
- Verificar la fidelidad de una afirmación contra el texto original requiere buscar manualmente entre decenas de páginas en inglés.
- Registrar citas formales con número de página y sección se vuelve una tarea manual propensa a errores.

Nuestra arquitectura resuelve esto convirtiendo el texto en una secuencia ordenada de **Segmentos Estructurados de Trazabilidad**.

---

## 2. Modelo de Datos del Segmento (`Segment`)

Cada párrafo o elemento atómico del paper se modela mediante la clase `Segment` definida en [`utils/file_handlers.py`](file:///c:/Users/Usuario/Documents/ghitub/translation_batch_app/utils/file_handlers.py):

```python
@dataclass
class Segment:
    id: int                       # Índice secuencial inmutable único (0..N)
    original: str                 # Texto exacto original en inglés
    translated: str = ""          # Traducción validada al español
    status: str = "pendiente"     # Estado: pendiente | traducido | ok | sospechoso | omitido | error
    validation_notes: List[str]   # Observaciones del Agente Validador
    retries: int = 0              # Conteo de reintentos automáticos
    color: str = "#FFD54F"        # Color asignado para sincronización visual

    # --- Metadatos de Trazabilidad Académica de Origen ---
    section: str = "General"      # Sección detectada (ej: "1. Introduction", "3. Methodology")
    subsection: str = ""         # Subsección (ej: "3.2 Data Collection")
    page: int = 1                # Número de página en el documento original (1-indexed)
    paragraph_num: int = 1       # Número de párrafo ordinal dentro de su sección o página
    element_type: str = "body"   # title | abstract | heading | body | caption | reference | formula
    is_marked: bool = False      # Bandera de selección/marcado por el usuario para citar o auditar
```

### Propiedades Calculadas de Procedencia
- **`provenance_label`**: 
  > `"Sección 3.2, página 7, párrafo 4 del documento original"`
- **`short_provenance`**:
  > `"Pág. 7 · P4 (3.2 Methodology)"`

---

## 3. Ciclo de Vida de los IDs y Trazabilidad a través del Pipeline

El ciclo de vida se ilustra en el siguiente diagrama de flujo:

```mermaid
graph TD
    A["Documento Original en Inglés (PDF / DOCX / TXT)"] --> B["Agente 1: Extractor Académico"]
    B -->|Genera N Segmentos con {id, section, page, paragraph_num, element_type}| C["Agente 2: Traductor Académico (Gemini Flash-Lite)"]
    C -->|Puebla .translated preservando citas y fórmulas| D["Agente 3: Validador Académico"]
    D -->|Valida integridad de citas [1] y autor-año; reintenta si hay fallas| E["Agente 4: Alineador Semántico"]
    E -->|Asigna color sincronizado y verifica coherencia semántica| F["Estado de Sesión (Streamlit Session State)"]
    F --> G["1. Descarga Inmediata (.docx, .pdf, .txt)"]
    F --> H["2. Interfaz Interactiva de Selección y Marcado"]
    H -->|El usuario selecciona párrafos| I["Párrafo Resaltado Visualmente + Ficha de Procedencia Exacta + Original Paralelo"]
    I --> J["Canasta de Citas (Copia en Markdown / Dossier)"]
    I --> K["3. Descarga Enriquecida (.docx y .pdf con resaltado y notas al pie)"]
```

### Paso a Paso del Ciclo de Vida:

1. **Extracción y Asignación de Coordenadas de Origen**:
   - En **PDF**: Se itera página a página (`page_idx = 1..P`). Cada bloque de texto se asigna a su página real y se detectan encabezados de sección y numeraciones.
   - En **DOCX**: Se detectan saltos de página explícitos (`w:br[@w:type="page"]`), estilos de encabezado (`Heading 1`, `Heading 2`, `Title`) y avance de volumen de palabras para calibrar páginas reales.
   - En **TXT**: Se identifican marcas como `--- Page X ---` y patrones de encabezados numerados.
   - Cada segmento recibe un `id` entero secuencial (0, 1, 2, ...), su `section`, su `page` y su `paragraph_num`.

2. **Inmutabilidad de los Metadatos en la Traducción**:
   - El **Agente Traductor** procesa el texto en lote párrafo por párrafo sin alterar los metadatos de procedencia (`page`, `section`, `paragraph_num`, `id`).
   - Los elementos detectados como `reference` o `formula` se marcan como omitidos de traducción automática para no corromper la notación matemática o los metadatos bibliográficos de autores y revistas.
   - Para los párrafos del cuerpo, el prompt académico instruye estrictamente conservar intactas todas las citas in-text como `[1]`, `[1, 2]` o `(Smith et al., 2020)`.

3. **Validación Automática de Integridad**:
   - El **Agente Validador** audita mediante expresiones regulares que ninguna cita entre corchetes o autor-año haya sido omitida o deformada durante la traducción. Si detecta discrepancias, dispara un reintento enfocado con backoff exponencial.

4. **Persistencia e Interacción en la Interfaz (Streamlit)**:
   - Los segmentos terminados residen en `st.session_state.results[filename]["segments"]`.
   - Cuando el usuario hace clic o selecciona un párrafo en el inspector o en el lienzo de lectura, se consulta su objeto `Segment` exacto por su `id`.
   - La interfaz despliega la tarjeta de procedencia:
     ```
     📍 Sección: 1. Introduction | 📄 Página: 1 | 🔢 Párrafo: 1 del documento original
     ```
   - Se muestra el texto original en inglés lado a lado con la traducción al español.
   - Al activar el marcado (`is_marked = True`), el párrafo queda registrado en la **Canasta de Citas**.

5. **Generación de Entregables Enriquecidos**:
   - **Descarga Inmediata**: Genera el documento completo traducido al español en `.docx` o `.pdf` listo para leer.
   - **Descarga Enriquecida**:
     - En `.docx`: Aplica resaltado en amarillo a los párrafos marcados y añade párrafos indentados en color índigo con la procedencia exacta: `📌 Procedencia: Sección X, página Y, párrafo Z del documento original`.
     - En `.pdf`: Dibuja una caja de resaltado suave (caja ámbar con borde redondeado) e inyecta la nota de procedencia directamente al pie del párrafo.
   - **Ficha de Citas (Dossier)**: Genera un resumen listo para copiar en Markdown con la traducción, el original en inglés y la cita formal de procedencia.

---

## 4. Comparativa de Métodos de Extracción

| Formato | Detección de Página | Detección de Sección | Manejo de Citas y Fórmulas |
| :--- | :--- | :--- | :--- |
| **PDF** | Nativa por objeto de página (`reader.pages`) | Regex de encabezados y mayúsculas sostenidas | Citas aisladas preservadas; fórmulas LaTeX detectadas |
| **DOCX** | Saltos de página XML + Estimación ponderada (~450 palabras/pág) | Estilos nativos de Word (`Heading 1`, `Heading 2`) | Párrafos nativos; fórmulas preservadas |
| **TXT** | Marcadores explícitos `--- Page X ---` o contador de palabras | Regex estructural de numeración (1. 2. 3.) | Texto plano; passthrough íntegro de citas |

---

## 5. Resumen de Ventajas para el Investigador

1. **Cero intervención manual**: La traducción completa arranca y concluye automáticamente con solo subir el archivo.
2. **Descarga inmediata**: Documento disponible al instante en formatos estándar (.docx / .pdf).
3. **Auditoría visual al clic**: Comprobación inmediata del inglés original con un solo clic sobre cualquier párrafo dudoso o citatorio.
4. **Exportación enriquecida**: Entregables con notas de origen incorporadas para facilitar la redacción de tesis y artículos científicos sin perder jamás la referencia original.
