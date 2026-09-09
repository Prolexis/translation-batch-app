"""
utils/file_handlers.py
------------------------
Lectura, extracción estructurada académica y exportación de documentos
(.txt, .docx, .pdf) con trazabilidad exacta de origen:
  - Sección y subsección
  - Número de página en el documento original
  - Número de párrafo dentro de la sección/página
  - Tipo de elemento (título, resumen, encabezado, cuerpo, cita, fórmula, referencia)
  - Marcado visual interactivo y exportación enriquecida
"""

import io
import re
import unicodedata
import logging
from dataclasses import dataclass, field
from typing import List, Tuple, Optional, Dict, Any

logger = logging.getLogger("translation_app.file_handlers")

# --- Importaciones de Librerías Externas con Resiliencia de Despliegue ---
try:
    import docx
    from docx import Document as DocxDocument
    try:
        from docx.enum.text import WD_COLOR_INDEX
    except ImportError:
        try:
            from docx.enum.text import WD_COLOR as WD_COLOR_INDEX
        except ImportError:
            WD_COLOR_INDEX = None
    from docx.shared import Pt, Inches, RGBColor
    DOCX_AVAILABLE = True
    _docx_error = ""
except Exception as _err:
    docx = None
    DocxDocument = None
    WD_COLOR_INDEX = None
    Pt = Inches = RGBColor = None
    DOCX_AVAILABLE = False
    _docx_error = str(_err)

try:
    from pypdf import PdfReader
    PYPDF_AVAILABLE = True
    _pypdf_error = ""
except Exception as _err:
    PdfReader = None
    PYPDF_AVAILABLE = False
    _pypdf_error = str(_err)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas
    try:
        from reportlab.lib.utils import simpleSplit
    except ImportError:
        def simpleSplit(text, fontName, fontSize, maxWidth):
            return text.split("\n")
    REPORTLAB_AVAILABLE = True
    _reportlab_error = ""
except Exception as _err:
    A4 = (595.27, 841.89)  # Fallback A4 dimensions in points
    cm = 28.3464567
    canvas = None
    simpleSplit = lambda text, *_: text.split("\n")
    REPORTLAB_AVAILABLE = False
    _reportlab_error = str(_err)


class UnsupportedFormatError(Exception):
    pass


class CorruptFileError(Exception):
    pass


@dataclass
class Segment:
    """
    Unidad atómica de texto que preserva la trazabilidad exacta de origen
    a lo largo de todo el pipeline de traducción.
    """
    id: int
    original: str
    translated: str = ""
    status: str = "pendiente"          # pendiente | traducido | ok | sospechoso | omitido | error
    validation_notes: List[str] = field(default_factory=list)
    retries: int = 0
    color: str = "#FFD54F"

    # Metadatos de Trazabilidad Académica
    section: str = "General"           # Ej: "1. Introduction", "3. Methodology", "References"
    subsection: str = ""              # Ej: "3.2 Data Collection"
    page: int = 1                     # Número de página en el documento original (1-indexed)
    paragraph_num: int = 1            # Número de párrafo dentro de la sección/página (1-indexed)
    element_type: str = "body"        # metadata | title | authors | affiliations | abstract | keywords | heading | body | caption | reference | formula
    is_marked: bool = False           # Marcado por el usuario para citación o auditoría

    @property
    def provenance_label(self) -> str:
        """Etiqueta formal de procedencia legible para el usuario."""
        parts = []
        if self.section and self.section != "General":
            parts.append(f"Sección {self.section}")
        if self.subsection:
            parts.append(f"Subsección {self.subsection}")
        parts.append(f"página {self.page}")
        parts.append(f"párrafo {self.paragraph_num}")
        return ", ".join(parts) + " del documento original"

    @property
    def short_provenance(self) -> str:
        """Etiqueta compacta para chips o insignias."""
        sec = self.section[:20] + "..." if len(self.section) > 20 else self.section
        return f"Pág. {self.page} · P{self.paragraph_num} ({sec})"


def repair_academic_symbols_and_ligatures(text: str) -> str:
    """Repara caracteres corruptos de PDF, ligaduras rotas, caracteres de dibujo de cajas y símbolos matemáticos."""
    if not text:
        return ""

    # 1. Ligaduras tipográficas estándar Unicode y caracteres invisibles/dañados
    ligatures = {
        "\ufb00": "ff", "\ufb01": "fi", "\ufb02": "fl", "\ufb03": "ffi", "\ufb04": "ffl",
        "\ufb05": "ft", "\ufb06": "st", "\xad": "", "\u200b": "", "\ufffd": "", "\ufeff": ""
    }
    for k, v in ligatures.items():
        text = text.replace(k, v)

    # Eliminar caracteres de control invisibles o corruptos
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ufeff\u200b\ufffd]', '', text)

    # 2. Reparar palabras cortadas por salto de línea con guión tipográfico
    text = re.sub(r'(\b[A-Za-z]+)-\s*\n\s*([a-z]+)', r'\1\2', text)
    text = re.sub(r'(\b[A-Za-z]{2,})-\s+([a-z]{2,}\b)', r'\1-\2', text)

    # 3. Corregir palabras típicas donde la ligadura 'fi'/'fl'/'ff' se extrajo como '■' o ''
    word_rep = [
        (r'\bsigni[■\ufffd]cantly\b', 'significantly'),
        (r'\bsigni[■\ufffd]cant\b', 'significant'),
        (r'\bveri[■\ufffd]cation\b', 'verification'),
        (r'\bveri[■\ufffd]caci[oó]n\b', 'verificación'),
        (r'\bOf[■\ufffd]ce\b', 'Office'),
        (r'\b[■\ufffd]nish\b', 'finish'),
        (r'\b[■\ufffd]nalizar\b', 'finalizar'),
        (r'\bde[■\ufffd]ne\b', 'define'),
        (r'\bde[■\ufffd]ned\b', 'defined'),
        (r'\bde[■\ufffd]nici[oó]n\b', 'definición'),
        (r'\bbene[■\ufffd]cio\b', 'beneficio'),
        (r'\be[■\ufffd]ciente\b', 'eficiente'),
        (r'\be[■\ufffd]cacia\b', 'eficacia'),
        (r'\bspeci[■\ufffd]c\b', 'specific'),
        (r'\bclasi[■\ufffd]caci[oó]n\b', 'clasificación'),
        (r'\bmodi[■\ufffd]caci[oó]n\b', 'modificación'),
        (r'\bidenti[■\ufffd]caci[oó]n\b', 'identificación'),
        (r'\bcon[■\ufffd]anza\b', 'confianza'),
        (r'\bcon[■\ufffd]guraci[oó]n\b', 'configuración'),
        (r'\bin[■\ufffd]uencia\b', 'influencia'),
        (r'\bin[■\ufffd]ujo\b', 'influjo'),
        (r'\bpro[■\ufffd]le\b', 'profile'),
        (r'\bdif[■\ufffd]cultad\b', 'dificultad'),
        (r'\baf[■\ufffd]liaci[oó]n\b', 'afiliación'),
        (r'\baf[■\ufffd]liaciones\b', 'afiliaciones'),
        (r'\bsuf[■\ufffd]ciente\b', 'suficiente'),
        (r'\bsuf[■\ufffd]cient\b', 'sufficient'),
        (r'\bcon[■\ufffd]icto\b', 'conflicto'),
        (r'\bcon[■\ufffd]ict\b', 'conflict'),
        (r'\bperfor[■\ufffd]', 'perfor-'),
        (r'\bsigni[■\ufffd]', 'signifi-'),
    ]
    for pat, rep in word_rep:
        text = re.sub(pat, rep, text, flags=re.I)

    # 4. Reemplazo de ligaduras genéricas residuales entre letras: ej. 'bene■cio' -> 'beneficio'
    text = re.sub(r'([A-Za-z])[■\ufffd]([a-z]{2,})', r'\1fi\2', text)

    # 5. Brackets y tuplas matemáticas (en IEEE se codifican \langle y \rangle que extraen como '■')
    text = re.sub(r'(?<=[=\s,:(])■(?=[A-Za-z0-9\\$])', '<', text)
    text = re.sub(r'(?<=[A-Za-z0-9\\$\)\]])■(?=[\s,.:;—\-\)]|$)', '>', text)
    text = re.sub(r'■\s*p\b', '|= p', text)

    # 6. Limpiar artefactos de dibujo de corchetes gigantes ASCII/Unicode (piezas de llave de ecuaciones)
    text = re.sub(r'[\u23a1-\u23b9\u2500-\u257f]+', ' ', text)

    # 7. Cualquier '■' remanente aislado se convierte en guión o espacio
    text = text.replace('■', '-')
    return text


def clean_math_display(text: str) -> str:
    """Convierte código LaTeX sin compilar y fórmulas en texto limpio y legible para visualización y PDF."""
    if not text:
        return ""

    # Casos por partes \begin{cases} ... \end{cases}
    def _cases_sub(m):
        content = m.group(1)
        raw_parts = [p.strip() for p in re.split(r'\\\\', content) if p.strip()]
        cleaned_parts = []
        for p in raw_parts:
            cp = re.sub(r'\s*&\s*', ' si ', p)
            cleaned_parts.append(cp)
        return '{ ' + ' ; '.join(cleaned_parts) + ' }'

    text = re.sub(r'\\begin\{cases\}([\s\S]*?)\\end\{cases\}', _cases_sub, text)
    # Fracciones: \frac{A}{B} -> (A)/(B)
    text = re.sub(r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'(\1)/(\2)', text)
    text = re.sub(r'\\frac\{([^{}]+)\}\{([^{}]+)\}', r'(\1)/(\2)', text)

    # Brackets y símbolos matemáticos comunes a Unicode matemático real
    text = text.replace(r'\langle', '<').replace(r'\rangle', '>')
    text = text.replace(r'\cdot', '·').replace(r'\times', '×')
    text = text.replace(r'\leq', '≤').replace(r'\geq', '≥')
    text = text.replace(r'\neq', '≠').replace(r'\approx', '≈')
    text = text.replace(r'\in', ' ∈ ').replace(r'\notin', ' ∉ ')
    text = text.replace(r'\mathbb{N}', 'ℕ').replace(r'\mathbb{R}', 'ℝ')
    text = text.replace(r'\perp', '⊥').replace(r'\top', '⊤')
    text = text.replace(r'\triangle', '△')
    text = text.replace(r'\leftarrow', '←').replace(r'\rightarrow', '→').replace(r'\to', '→')
    text = text.replace(r'\max', 'max').replace(r'\min', 'min').replace(r'\arg', 'arg')
    text = text.replace(r'\sum', '∑').replace(r'\prod', '∏')
    text = text.replace(r'\dots', '…').replace(r'\cdots', '…')

    text = re.sub(r'\\text\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\mathrm\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\mathbf\{([^{}]+)\}', r'\1', text)

    # Letras griegas LaTeX en símbolos Unicode reales
    greek_latex = {
        r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
        r'\epsilon': 'ε', r'\varepsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η',
        r'\theta': 'θ', r'\iota': 'ι', r'\kappa': 'κ', r'\lambda': 'λ',
        r'\mu': 'μ', r'\nu': 'ν', r'\xi': 'ξ', r'\pi': 'π', r'\rho': 'ρ',
        r'\sigma': 'σ', r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ',
        r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
        r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ',
        r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Phi': 'Φ',
        r'\Psi': 'Ψ', r'\Omega': 'Ω'
    }
    for g_cmd, g_val in greek_latex.items():
        text = re.sub(re.escape(g_cmd) + r'(?![a-zA-Z])', g_val, text)

    text = re.sub(r'_\{([^{}]+)\}', r'_\1', text)
    text = re.sub(r'\^\{([^{}]+)\}', r'^\1', text)
    text = text.replace('$', '')
    return text


def _is_running_header_or_footer(line: str) -> bool:
    """Filtra líneas de encabezado editorial, volúmenes de revista, números de página y avisos legales."""
    clean = line.strip()
    if not clean:
        return True
    # 1. Indicadores de volumen y año IEEE (ej. "VOLUME 9, 2021 102717", "VOLUMEN 9, 2021")
    if re.search(r'\b(?:VOLUME|VOLUMEN)\s+\d+(?:,\s*\d{4})?(?:\s+\d+)?\b', clean, re.I):
        return True
    if re.search(r'^\d{5,8}\s+(?:VOLUME|VOLUMEN)\b', clean, re.I):
        return True
    if re.search(r'^(?:WORKLOAD\s+)?(?:VOLUME|VOLUMEN)\s+\d+', clean, re.I):
        return True
    # 2. Avisos editoriales, licencias y copyright
    if re.search(r'(?:associate editor coordinating|approving it for publication|creative commons attribution|licensed under a creative commons|this work is licensed|date of current version|digital object identifier|all rights reserved)', clean, re.I):
        return True
    if re.search(r'^(?:See\s+https?://|https?://creativecommons\.org)', clean, re.I):
        return True
    # 3. Encabezados repetidos de página con título truncado
    if re.search(r'^(?:MODELADO|SURVEY|IEEE\s+ACCESS|PROCEEDINGS|TRANSACTIONS|REVISTA).*P[aá]g\.?\s*\d+$', clean, re.I):
        return True
    # 4. Marcadores de página aislados
    if re.search(r'^(?:—|-|–)\s*P[aá]gina\s*\d+\s*(?:—|-|–)$', clean, re.I):
        return True
    if re.search(r'^(?:Page|Página|Pág\.)\s*\d+(?:\s*(?:of|de)\s*\d+)?$', clean, re.I):
        return True
    # 5. Números de artículo IEEE aislados (ej. 102716)
    if re.match(r'^\d{5,8}$', clean):
        return True
    # 6. Encabezados de revistas académicas (ej. "Transportation Research Part E 186 (2024) 103563", "Computers in Biology... 166 (2023)")
    if re.search(r'^(?:[A-Za-z\s\:\.\,\&\-]+\s+)?\d+\s*\(\d{4}\)\s*\d+$', clean):
        return True
    # 7. Autores en encabezado/pie de página (ej. "I. Abdulrashid et al.", "I. Abdulrashid y otros", "S. Ali et al.")
    if re.match(r'^[A-Z]\.\s+[A-Za-z\-]+(?:\s+(?:et\s+al\.?|y\s+(?:otros|cols\.?|\bals?\b)))$', clean):
        return True
    return False


def _clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = repair_academic_symbols_and_ligatures(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# Patrones de encabezados y secciones académicas (secciones numéricas limitadas a 1..29 o romanos con punto/espacio, excluyendo años)
_SECTION_REGEX = re.compile(
    r"^(?:(?:(?:[1-9]|1\d|2\d)(?:\.\d+)*(?:\.\s*|\s+)|(?:I|II|III|IV|V|VI|VII|VIII|IX|X|XI|XII)(?:\.\s*|\s+))([A-Za-z\u00C0-\u017F\s\-]{2,65})|"
    r"(abstract|resumen|introduction|introducci[oó]n|background|related work|"
    r"methods|methodology|metodolog[ií]a|materials and methods|results|resultados|"
    r"discussion|discusi[oó]n|conclusion|conclusions|conclusiones|"
    r"acknowledgments|agradecimientos|references|bibliography|bibliograf[ií]a))(?::|\.|\b)?$",
    re.IGNORECASE
)

_REFERENCES_HEADER_REGEX = re.compile(
    r"^(?:(\d+\.?\s+)?(references|bibliography|bibliograf[ií]a|referencias\s+bibliogr[aá]ficas))(?::|\.|\b)?$",
    re.IGNORECASE
)

_MATH_FORMULA_REGEX = re.compile(
    r"(\$\$[\s\S]+?\$\$|\\\[[\s\S]+?\\\]|\\begin\{(?:equation|align|gather)\*?\}[\s\S]+?\\end\{(?:equation|align|gather)\*?\})"
)


def highlight_citations_html(text: str, is_marked: bool = False) -> str:
    """Resalta visualmente con color y subrayado las citas in-text [1], autor-año (Author, Year) y fragmentos exactos <mark>...</mark> o <u>...</u>."""
    if not text:
        return ""
    # Reparar ligaduras y limpiar expresiones matemáticas para lectura fluida
    text = repair_academic_symbols_and_ligatures(text)
    text = clean_math_display(text)

    import html as _html
    safe = _html.escape(text)

    # Convertir &lt;mark&gt; y &lt;u&gt; a subrayado y resaltado visual inline exacto
    underline_style = (
        "text-decoration: underline 2.5px #D97706; text-underline-offset: 3.5px; "
        "background-color: rgba(254, 240, 138, 0.55); color: inherit; "
        "padding: 1px 3px; border-radius: 2px; font-weight: 600; "
        "box-decoration-break: clone; -webkit-box-decoration-break: clone;"
    )

    safe = re.sub(
        r"&lt;(?:mark|u)&gt;([\s\S]*?)&lt;/(?:mark|u)&gt;",
        rf'<span class="academic-inline-highlight" style="{underline_style}">\1</span>',
        safe
    )

    # Pastilla de cita adaptable a tema claro y oscuro
    badge_style = "color: var(--citation-normal-color, #3730A3); background: var(--citation-normal-bg, #EEF2FF); font-weight: 700; padding: 2px 6px; border-radius: 4px; border: 1px solid var(--citation-normal-border, #818CF8);"
    badge_class = "academic-citation"

    def _rep(m):
        return f'<span class="{badge_class}" style="{badge_style}">{m.group(0)}</span>'

    # 1. Citas entre corchetes [1], [2, 3], [4-6]
    safe = re.sub(r"\[\s*\d+(?:[\s,\-–—]+\d+)*\s*\]", _rep, safe)

    # 2. Citas autor-año (Vaswani et al., 2017) o (Devlin, 2019)
    safe = re.sub(r"\([A-Z][a-zA-Z\s\.\&]+(?:et\s+al\.?)?,?\s*(?:19|20)\d{2}[a-z]?\)", _rep, safe)

    return safe


def _detect_element_type(text: str, current_section: str, page_num: int, is_first: bool = False) -> Tuple[str, Optional[str], Optional[str]]:
    """
    Detecta el tipo de elemento académico y actualiza sección/subsección si aplica.
    Retorna: (element_type, new_section, new_subsection)
    """
    clean_line = text.strip()

    # Detectar si es fórmula matemática pura
    if _MATH_FORMULA_REGEX.search(clean_line) and len(clean_line) < 300:
        return ("formula", None, None)
    if clean_line.startswith(("$$", "\\[", "\\begin{equation}")):
        return ("formula", None, None)

    # Detectar si estamos en sección de Referencias/Bibliografía
    if _REFERENCES_HEADER_REGEX.match(clean_line):
        return ("heading", clean_line, "")
    if current_section and _REFERENCES_HEADER_REGEX.match(current_section):
        return ("reference", None, None)

    # Título en página 1
    if page_num == 1 and is_first and len(clean_line) < 200 and not clean_line.endswith("."):
        return ("title", "Title", "")

    # Abstract
    if re.match(r'^(?:A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|abstract|R\s*E\s*S\s*U\s*M\s*E\s*N|resumen)\b', clean_line, re.I):
        return ("abstract", "Abstract", "")

    # Detección de encabezados o secciones numeradas (con filtros anti-falsos positivos)
    words = clean_line.split()
    last_word = words[-1].lower() if words else ""
    is_continuation = (
        last_word in {"through", "and", "of", "the", "to", "in", "by", "with", "for", "from", "are", "is", "that", "which"}
        or len(words) > 10
        or clean_line.endswith((",", ";", "..."))
    )

    if not is_continuation:
        m = _SECTION_REGEX.match(clean_line)
        if m and len(clean_line) < 80:
            num = m.group(1)
            if num and "." in num and not num.endswith("."):
                # Subsección, ej: 3.2 Methods
                return ("heading", None, clean_line)
            return ("heading", clean_line, "")

    return ("body", None, None)


# --------------------------------------------------------------------------
# Extractores con Trazabilidad (PDF, DOCX, TXT)
# --------------------------------------------------------------------------

def _clamp_paragraph(text: str, max_chars: int = 850) -> List[str]:
    """Divide párrafos excesivamente largos en oraciones gramaticales legibles."""
    text = text.strip()
    if len(text) <= max_chars:
        return [text]
    sentences = re.split(r'(?<=\.)\s+(?=[A-Z\u00C0-\u017F0-9])', text)
    result = []
    curr = []
    curr_len = 0
    for s in sentences:
        if curr and (curr_len + len(s) + 1 > max_chars):
            result.append(" ".join(curr))
            curr = [s]
            curr_len = len(s)
        else:
            curr.append(s)
            curr_len += len(s) + 1
    if curr:
        result.append(" ".join(curr))
    return result if result else [text]


def _decompose_page1_academic(lines: List[str]) -> List[Tuple[str, Optional[str], str]]:
    """
    Descompone rigurosamente la primera página de un paper científico (IEEE, Elsevier, ACM, Springer, etc.)
    en sus componentes estructurales reales:
      (element_type, section_name, text)
      Tipos: metadata, title, authors, affiliations, abstract, keywords, heading, body, table, caption
    """
    # 1. Preprocesar líneas: separar encabezados estructurales pegados o embebidos en la misma línea
    # (común en papers con layouts de 2 columnas o barras laterales, ej: "HealthcareA B S T R A C T")
    preprocessed: List[str] = []
    embedded_header_pat = re.compile(
        r'(?:^|(?<=[a-z0-9\s]))('
        r'A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|ABSTRACT|R\s*E\s*S\s*U\s*M\s*E\s*N|RESUMEN|'
        r'K\s*E\s*Y\s*W\s*O\s*R\s*D\s*S|KEYWORDS|INDEX\s+TERMS|PALABRAS\s+CLAVE|'
        r'A\s*R\s*T\s*I\s*C\s*L\s*E\s*I\s*N\s*F\s*O|ARTICLE\s+INFO'
        r')(?:[\:\—\-\.\s]|$)',
        re.I
    )

    for line in lines:
        # Separar avisos de copyright / metadatos pegados directamente al inicio del título
        m_meta = re.search(r'^(.*?all rights reserved\.?)\s*(?=[A-Z])', line, re.I)
        if m_meta:
            preprocessed.append(m_meta.group(1).strip())
            line = line[m_meta.end():].strip()

        m = embedded_header_pat.search(line)
        if m and m.start() > 0:
            prefix = line[:m.start()].strip()
            rest = line[m.start():].strip()
            if prefix:
                preprocessed.append(prefix)
            if rest:
                preprocessed.append(rest)
            continue
        preprocessed.append(line)

    results: List[Tuple[str, Optional[str], str]] = []
    state = "SEARCH_METADATA"
    curr_text: List[str] = []
    current_sec_label: Optional[str] = None

    def flush(target_state: str, sec_label: Optional[str] = None):
        nonlocal curr_text
        if curr_text:
            text = " ".join(curr_text).strip()
            if text:
                t = target_state.lower()
                results.append((t, sec_label, text))
            curr_text = []

    for line in preprocessed:
        is_meta_marker = bool(re.search(
            r'(?:digital object identifier|doi[\:\s]|10\.\d{4,9}/|received\s+[a-z0-9]+|recibido\s+|accepted\s+|aceptado\s+|date of publication|fecha de publicaci[oó]n|current version|available online|disponible en l[ií]nea|contents lists available at|sciencedirect|journal homepage|homepage[\:\s]|www\.|https?://|published by|elsevier|springer|wiley|ieee\s+access|creative commons|open access|cc\s+by|\b\d{4}-\d{3}[\dXx]\b|\b\d+\s*\(\d{4}\)\s*\d+)',
            line, re.I
        ))
        is_article_info = bool(re.search(r'^(?:A\s*R\s*T\s*I\s*C\s*L\s*E\s*I\s*N\s*F\s*O|ARTICLE\s+INFO|ARTICLE\s+HISTORY)', line, re.I))
        is_abs_start = bool(re.search(r'^(?:A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|ABSTRACT|R\s*E\s*S\s*U\s*M\s*E\s*N|RESUMEN)(?:[\:\—\-\.\s]|$)', line, re.I))
        is_key_start = bool(re.search(r'^(?:K\s*E\s*Y\s*W\s*O\s*R\s*D\s*S|KEYWORDS|INDEX\s+TERMS|PALABRAS\s+CLAVE|KEY\s+WORDS|T[EÉ]RMINOS\s+DE\s+[IÍ]NDICE)(?:[\:\—\-\.\s]|$)', line, re.I))
        is_sec_head = bool(re.match(r'^(?:(?:[IVXLCDM]+|[1-9]\d?)\.|\d+\.\d+)\s*[A-Z\u00C0-\u017F]', line))
        is_affil_marker = bool(
            re.search(r'^[a-z0-9\*\†\§\d\s\.\,\-]*(?:department|departamento|faculty|facultad|school|escuela|university|universidad|institute|instituto|laboratory|laboratorio|center|centre|division|corresponding author|autor de correspondencia|this work was supported|este trabajo fue financiado|email[\:\s])', line, re.I)
        )
        is_author_marker = bool(
            re.search(r'\b(?:dr\.|prof\.|phd|member|fellow|senior member),?\s*(?:ieee)?\b', line, re.I)
            or re.search(r'\b(?:and|y)\s+[A-Z\s\.\-]{3,}\b', line)
            or re.search(r'^[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*(?:[a-z0-9\*\†\§\,]+)?(?:\s*,\s*[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*(?:[a-z0-9\*\†\§\,]+)?)+', line)
            or (line.isupper() and len(line.split()) >= 2 and len(line.split()) <= 12 and not line.endswith('.'))
        )

        if is_article_info:
            if curr_text:
                flush(state, "Afiliaciones" if state == "AFFILIATIONS" else None)
            state = "METADATA"
            continue

        if is_abs_start:
            flush(state, "Palabras Clave" if state == "KEYWORDS" else "Afiliaciones" if state == "AFFILIATIONS" else "Título" if state == "TITLE" else None)
            state = "ABSTRACT"
            clean = re.sub(r'^(?:A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|ABSTRACT|R\s*E\s*S\s*U\s*M\s*E\s*N|RESUMEN)\s*[\:\—\-\.]*\s*', '', line, flags=re.I).strip()
            if clean:
                curr_text.append(clean)
            continue

        if is_key_start:
            flush(state, "Resumen" if state == "ABSTRACT" else "Afiliaciones" if state == "AFFILIATIONS" else "Título" if state == "TITLE" else None)
            state = "KEYWORDS"
            clean = re.sub(r'^(?:K\s*E\s*Y\s*W\s*O\s*R\s*D\s*S|KEYWORDS|INDEX\s+TERMS|PALABRAS\s+CLAVE|KEY\s+WORDS|T[EÉ]RMINOS\s+DE\s+[IÍ]NDICE)\s*[\:\—\-\.]*\s*', '', line, flags=re.I).strip()
            if clean:
                curr_text.append(clean)
            continue

        if is_sec_head:
            flush(state, "Resumen" if state == "ABSTRACT" else "Palabras Clave" if state == "KEYWORDS" else "Afiliaciones" if state == "AFFILIATIONS" else None)
            results.append(("heading", line, line))
            current_sec_label = line
            state = "BODY"
            continue

        if state == "SEARCH_METADATA":
            if is_meta_marker:
                curr_text.append(line)
                state = "METADATA"
                continue
            else:
                state = "TITLE"
                curr_text.append(line)
                continue

        if state == "METADATA":
            if is_meta_marker:
                curr_text.append(line)
                continue
            else:
                flush("METADATA", "Encabezado")
                state = "TITLE"
                curr_text.append(line)
                continue

        if state == "TITLE":
            if is_meta_marker:
                curr_text.append(line)
                flush("METADATA", "Encabezado")
                state = "TITLE"
                continue
            elif is_affil_marker:
                flush("TITLE", "Título")
                state = "AFFILIATIONS"
                curr_text.append(line)
                continue
            elif is_author_marker:
                flush("TITLE", "Título")
                state = "AUTHORS"
                curr_text.append(line)
                continue
            else:
                curr_text.append(line)
                continue

        if state == "AUTHORS":
            if is_affil_marker:
                flush("AUTHORS", "Autores")
                state = "AFFILIATIONS"
                curr_text.append(line)
                continue
            else:
                curr_text.append(line)
                continue

        if state == "AFFILIATIONS":
            curr_text.append(line)
            continue

        if state == "BODY":
            # Si en el cuerpo de página 1 aparece metadato al pie (autor de correspondencia, correo, DOI, recibido)
            # flushear inmediatamente el cuerpo para evitar que la nota al pie se fugue dentro de la Introducción
            is_footnote_marker = bool(re.search(
                r'(?:^\*|\bcorresponding author\b|\bautor de correspondencia\b|\be-?mail addresses?\b|\bcontents lists available at\b|\bjournal homepage\b|https?://doi\.org/|\breceived\s+\d+|\brecibido\s+\d+)',
                line, re.I
            ))
            if is_footnote_marker:
                flush("BODY", current_sec_label or "Introducción")
                state = "METADATA"
                curr_text.append(line)
                continue

        curr_text.append(line)

    if curr_text:
        flush(state, "Resumen" if state == "ABSTRACT" else "Palabras Clave" if state == "KEYWORDS" else current_sec_label if state == "BODY" else "Encabezado")

    return results


def _is_table_start(line: str) -> bool:
    clean = line.strip()
    return bool(re.match(r'^(?:Table|Tabla)\s+\d+(?:[\.\:\—\-\s]+(?!(?:shows|presents|illustrates|depicts|indicates|displays|is|was|were|muestra|presenta|ilustra)\b)[A-Z\u00C0-\u017F0-9].*|$)', clean, re.I))


def _is_fig_caption(line: str) -> bool:
    clean = line.strip()
    return bool(re.match(r'^(?:Fig\.|Figure|Figura)\s*\d+[\.\:\—\-]', clean, re.I))


def _is_prose_sentence(line: str) -> bool:
    clean = line.strip()
    if len(clean) < 70:
        return False
    if '%' in clean or '√' in clean or '✓' in clean:
        return False
    if re.search(r'\b(?:accuracy|sensitivity|specificity|auc|f-measure|g-mean|class 0|class 1|class 2|pcrash|seat_im|trav_sp|vnum_lan|vspd_lim)\b', clean, re.I):
        return False
    if clean.endswith(('.', ':', ';')) and re.match(r'^[A-Z]', clean):
        words = clean.split()
        if len(words) >= 12:
            return True
    return False


def _parse_body_page_blocks(lines: List[str], current_section: str) -> List[Tuple[str, Optional[str], str]]:
    """Procesa páginas agrupando párrafos y detectando encabezados, tablas estructuradas, figuras y referencias."""
    results: List[Tuple[str, Optional[str], str]] = []
    curr: List[str] = []

    def flush_curr():
        nonlocal curr
        if curr:
            clamped = _clamp_paragraph(" ".join(curr))
            for cp in clamped:
                results.append(("body", current_section, cp))
            curr = []

    i = 0
    while i < len(lines):
        line = lines[i]

        # 1. Separar encabezado de tabla pegado al final de una oración
        m_emb_table = re.search(r'(?<=\s)(?:Table|Tabla)\s+\d+\s*$', line, re.I)
        if m_emb_table:
            prefix = line[:m_emb_table.start()].strip()
            if prefix:
                curr.append(prefix)
            line = line[m_emb_table.start():].strip()

        # 2. Detección de tablas completas
        if _is_table_start(line):
            flush_curr()
            t_lines = [line]
            i += 1
            while i < len(lines):
                cur_l = lines[i]
                if _SECTION_REGEX.match(cur_l) or _REFERENCES_HEADER_REGEX.match(cur_l) or _is_fig_caption(cur_l) or _is_table_start(cur_l):
                    break
                if _is_prose_sentence(cur_l):
                    break
                t_lines.append(cur_l)
                i += 1
            results.append(("table", current_section, "\n".join(t_lines)))
            continue

        # 3. Detección de pies de figura (captions)
        if _is_fig_caption(line):
            flush_curr()
            c_lines = [line]
            i += 1
            while i < len(lines) and not lines[i].endswith('.') and len(lines[i]) < 110:
                cur_l = lines[i]
                if _SECTION_REGEX.match(cur_l) or _is_table_start(cur_l) or _is_fig_caption(cur_l):
                    break
                c_lines.append(cur_l)
                i += 1
            results.append(("caption", current_section, " ".join(c_lines)))
            continue

        # 4. Encabezados, referencias y viñetas
        is_heading = bool(_SECTION_REGEX.match(line) or _REFERENCES_HEADER_REGEX.match(line))
        is_ref_item = bool(re.match(r"^\[\d+\]", line))
        is_bullet = bool(re.match(r"^[•\-\*]\s+", line))

        if is_heading or is_ref_item or is_bullet:
            flush_curr()
            if is_heading:
                results.append(("heading", line, line))
            elif is_ref_item:
                results.append(("reference", "Referencias", line))
            i += 1
            continue

        # 5. Segmentación de párrafos justificadas
        if curr:
            prev_line = curr[-1]
            is_terminator = prev_line.endswith((".", ":", "?", "!"))
            is_paragraph_end = is_terminator and (len(prev_line) < 48 or len(curr) >= 14 or line.startswith("   "))
            if is_paragraph_end and (line[0].isupper() or line.startswith(("$$\\", "$", "["))):
                flush_curr()
                curr = [line]
                i += 1
                continue

        curr.append(line)
        i += 1

    flush_curr()
    return results


def _char_width(c: str, font_size: float) -> float:
    """Estimación tipográfica proporcional del ancho de un carácter según su anatomía."""
    if c in "ijlItfr.,:;!|()[]-/'`\"":
        return font_size * 0.30
    if c in "mwMW@%&":
        return font_size * 0.80
    if c.isupper():
        return font_size * 0.65
    return font_size * 0.50


def _filter_text_in_horizontal_box(text: str, x_start: float, font_size: float, bx0: float, bx1: float) -> str:
    """Filtra con precisión tipográfica las palabras de un fragmento de texto que caen dentro de [bx0, bx1]."""
    text = text.strip()
    if not text:
        return ""
    total_w = sum(_char_width(c, font_size) for c in text)
    x_end = x_start + total_w

    if x_start >= (bx0 - 2) and x_end <= (bx1 + 2):
        return text
    if x_end < (bx0 - 2) or x_start > (bx1 + 2):
        return ""

    words = text.split(" ")
    matched_words = []
    curr_x = x_start
    for w in words:
        if not w:
            curr_x += _char_width(" ", font_size)
            continue
        w_w = sum(_char_width(c, font_size) for c in w)
        w_mid = (curr_x + curr_x + w_w) / 2.0
        # La palabra pertenece a la anotación solo si su punto medio está estrictamente dentro de los límites
        if (bx0 - 2) <= w_mid <= (bx1 + 2):
            matched_words.append(w)
        curr_x += w_w + _char_width(" ", font_size)

    return " ".join(matched_words)


def _extract_pdf_page_highlights(page) -> List[Dict[str, Any]]:
    """
    Extrae fragmentos de texto anotados con resaltado (/Highlight) o subrayado (/Underline, /Squiggly)
    en una página de PDF. Retorna lista de diccionarios con {'text': str, 'lines': List[str], 'color': str}.
    Maneja referencias indirectas de /Annots, agrupa quads en orden de lectura y repara saltos con guión.
    """
    highlights: List[Dict[str, Any]] = []
    annots = getattr(page, "annotations", None)
    if not annots:
        annots = page.get("/Annots")
    if hasattr(annots, "get_object"):
        try:
            annots = annots.get_object()
        except Exception:
            pass
    if not annots:
        return highlights

    annot_items: List[Dict[str, Any]] = []
    for a in annots:
        try:
            obj = a.get_object() if hasattr(a, "get_object") else a
            if hasattr(obj, "get_object"):
                obj = obj.get_object()
            if not isinstance(obj, dict):
                continue
            subtype = str(obj.get("/Subtype", "")).lstrip("/")
            if subtype.lower() not in ("highlight", "underline", "squiggly", "strikeout"):
                continue

            # Extraer color si existe
            c_val = obj.get("/C")
            color_hex = "#FDE047"  # Amarillo por defecto
            if c_val and hasattr(c_val, "__iter__"):
                try:
                    rgb = [max(0.0, min(1.0, float(x))) for x in c_val]
                    if len(rgb) >= 3:
                        r, g, b = int(rgb[0] * 255), int(rgb[1] * 255), int(rgb[2] * 255)
                        color_hex = f"#{r:02X}{g:02X}{b:02X}"
                except Exception:
                    pass

            # Coordenadas agrupadas por anotación: QuadPoints o Rect
            annot_boxes = []
            quads = obj.get("/QuadPoints")
            if hasattr(quads, "get_object"):
                try:
                    quads = quads.get_object()
                except Exception:
                    pass

            if quads and hasattr(quads, "__iter__"):
                q_coords = [float(x) for x in quads]
                for i in range(0, len(q_coords) - 7, 8):
                    xs = q_coords[i:i+8:2]
                    ys = q_coords[i+1:i+8:2]
                    y_min, y_max = min(ys), max(ys)
                    if subtype.lower() in ("underline", "squiggly") or (y_max - y_min) < 6:
                        # Para subrayado la línea está en la base (y_min). El texto asciende.
                        # NO restar hacia abajo para evitar fugar a la siguiente línea del paper.
                        by0 = y_min - 0.5
                        by1 = max(y_max, y_min + 11.5)
                    else:
                        by0 = y_min - 0.5
                        by1 = y_max + 0.5
                    # Cada quad de QuadPoints representa exactamente un renglón
                    annot_boxes.append((min(xs), by0, max(xs), by1, True))
            else:
                rect = obj.get("/Rect")
                if hasattr(rect, "get_object"):
                    try:
                        rect = rect.get_object()
                    except Exception:
                        pass
                if rect and hasattr(rect, "__iter__") and len(rect) >= 4:
                    r = [float(x) for x in rect]
                    bx0, by0, bx1, by1 = min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3])
                    is_single = (by1 - by0) <= 15.0
                    if subtype.lower() in ("underline", "squiggly") or (by1 - by0) < 6:
                        by0 = by0 - 0.5
                        by1 = max(by1, by0 + 11.5)
                    annot_boxes.append((bx0, by0, bx1, by1, is_single))

            if annot_boxes:
                # Agrupar en bandas de línea (tolerancia 5 pt) y ordenar de arriba hacia abajo, izquierda a derecha
                annot_boxes.sort(key=lambda b: (-round(((b[1] + b[3]) / 2.0) / 5.0), b[0]))
                annot_items.append({"boxes": annot_boxes, "color": color_hex, "obj": obj})
        except Exception as e:
            logger.debug("Error procesando anotación de PDF: %s", e)

    if not annot_items:
        return highlights

    # Mapeo (annot_idx, box_idx) -> fragmentos de texto
    box_texts: Dict[Tuple[int, int], List[str]] = {}
    for a_idx, item in enumerate(annot_items):
        for b_idx in range(len(item["boxes"])):
            box_texts[(a_idx, b_idx)] = []

    def visitor(text, cm, tm, font_dict, font_size):
        if not text or not text.strip():
            return
        try:
            if cm:
                x_start = cm[0] * tm[4] + cm[2] * tm[5] + cm[4]
                y = cm[1] * tm[4] + cm[3] * tm[5] + cm[5]
                fsize = (font_size if font_size and font_size > 0 else 10.0) * abs(cm[3])
            else:
                x_start = tm[4]
                y = tm[5]
                fsize = font_size if font_size and font_size > 0 else 10.0
            y_center_text = y + (fsize * 0.35)

            for a_idx, item in enumerate(annot_items):
                for b_idx, box_info in enumerate(item["boxes"]):
                    bx0, by0, bx1, by1 = box_info[0], box_info[1], box_info[2], box_info[3]
                    is_single = box_info[4] if len(box_info) > 4 else True

                    if is_single:
                        # Caja de renglón individual: aislamiento vertical estricto (rechaza líneas adyacentes)
                        y_center_box = (by0 + by1) / 2.0
                        v_match = abs(y_center_text - y_center_box) <= (fsize * 0.55)
                    else:
                        # Bloque multi-línea
                        v_match = (by0 <= y_center_text <= by1)

                    if v_match:
                        matched_snippet = _filter_text_in_horizontal_box(text, x_start, fsize, bx0, bx1)
                        if matched_snippet:
                            box_texts[(a_idx, b_idx)].append(matched_snippet)
        except Exception:
            pass

    try:
        page.extract_text(visitor_text=visitor)
    except Exception as e:
        logger.debug("Error en visitor_text de PDF: %s", e)

    for a_idx, item in enumerate(annot_items):
        lines_text = []
        for b_idx in range(len(item["boxes"])):
            line_str = " ".join(box_texts.get((a_idx, b_idx), [])).strip()
            line_str = re.sub(r"\s+", " ", line_str)
            if line_str:
                lines_text.append(line_str)

        full_annot_text = ""
        if lines_text:
            for line in lines_text:
                line = line.strip()
                if not line:
                    continue
                if not full_annot_text:
                    full_annot_text = line
                else:
                    if full_annot_text.rstrip().endswith("-"):
                        full_annot_text = full_annot_text.rstrip()[:-1] + line.lstrip()
                    else:
                        full_annot_text = full_annot_text + " " + line

            full_annot_text = repair_academic_symbols_and_ligatures(full_annot_text)
            full_annot_text = re.sub(r"\s+", " ", full_annot_text).strip()

        # Si no se extrajo por coordenadas o fue muy corto, fallback a /Contents
        if not full_annot_text or len(full_annot_text) < 3:
            contents_text = str(item.get("obj", {}).get("/Contents", "") or "").strip()
            if contents_text and len(contents_text) >= 3:
                full_annot_text = repair_academic_symbols_and_ligatures(contents_text)
                full_annot_text = re.sub(r"\s+", " ", full_annot_text).strip()
                lines_text = [full_annot_text]

        if full_annot_text and len(full_annot_text) >= 2:
            cleaned_lines = [
                repair_academic_symbols_and_ligatures(l).strip()
                for l in lines_text if len(l.strip()) >= 2
            ]
            highlights.append({
                "text": full_annot_text,
                "lines": cleaned_lines,
                "color": item["color"],
            })

    return highlights


def _build_flexible_phrase_regex(phrase: str) -> str:
    """
    Construye un patrón regex estricto pero tolerante a saltos de línea con guión,
    espacios múltiples, ligaduras y corchetes de citas bibliográficas como [79].
    """
    phrase = repair_academic_symbols_and_ligatures(phrase)
    tokens = re.findall(r'[a-zA-Z0-9]+|[^\s\w]', phrase)
    if not tokens:
        return ""
    pats = []
    for i, t in enumerate(tokens):
        if t in "-—–":
            pats.append(r'[-\u2010-\u2015]?\s*')
        elif t in '()[]{}':
            pats.append(re.escape(t) + r'\s*')
        elif t in '.,;:':
            if i == len(tokens) - 1:
                pats.append(re.escape(t))
            else:
                pats.append(re.escape(t) + r'\s*')
        else:
            if i < len(tokens) - 1 and tokens[i+1] not in '-—–.,;:)]}':
                pats.append(re.escape(t) + r'\s+')
            else:
                pats.append(re.escape(t) + r'\s*')
    return r''.join(pats)


def _inject_inline_marks(text_val: str, page_highlights: List[Any]) -> Tuple[str, bool, str]:
    """
    Inserta etiquetas <mark>...</mark> alrededor de las frases exactas detectadas
    como subrayadas/resaltadas en el PDF original.
    Soporta frases completas multi-línea, fragmentos de frases que cruzan límites de párrafo,
    y múltiples anotaciones en una misma página con estricto respeto de fronteras.
    Retorna: (texto_con_marcas, tiene_marcas, color_hex)
    """
    if not text_val or not page_highlights:
        return text_val, False, "#FFD54F"

    marked_text = text_val
    has_any_mark = False
    chosen_color = "#FFD54F"

    for item in page_highlights:
        if isinstance(item, dict):
            if item.get("consumed"):
                continue
            hl_text = item.get("text", "")
            lines = item.get("lines", [hl_text])
            color_hex = item.get("color", "#FFD54F")
        elif isinstance(item, (tuple, list)):
            hl_text = item[0]
            lines = [hl_text]
            color_hex = item[1] if len(item) > 1 else "#FFD54F"
        else:
            continue

        if not hl_text or len(hl_text.strip()) < 3:
            continue

        # 1. Normalizar ligaduras y guiones de salto de línea sin mutilar palabras compuestas
        clean_hl = repair_academic_symbols_and_ligatures(hl_text)
        clean_hl = re.sub(r"(\w+)-\s*\n\s*(\w+)", r"\1\2", clean_hl)
        clean_hl = re.sub(r"\s+", " ", clean_hl).strip()

        if len(clean_hl) < 3:
            continue

        # Si ya está completamente dentro de una marca en este texto, marcar consumido y continuar
        if f"<mark>{clean_hl}</mark>" in marked_text:
            has_any_mark = True
            chosen_color = color_hex
            if isinstance(item, dict):
                item["consumed"] = True
            continue

        # 2. Búsqueda exacta directa de la frase completa
        if clean_hl in marked_text and f"<mark>{clean_hl}</mark>" not in marked_text:
            marked_text = marked_text.replace(clean_hl, f"<mark>{clean_hl}</mark>", 1)
            has_any_mark = True
            chosen_color = color_hex
            if isinstance(item, dict):
                item["consumed"] = True
            continue

        # 3. Búsqueda flexible tolerante a guiones y espacios
        flex_pat = _build_flexible_phrase_regex(clean_hl)
        if flex_pat:
            m_flex = re.search(flex_pat, marked_text, flags=re.IGNORECASE)
            if m_flex:
                matched_span = m_flex.group(0).rstrip()
                marked_text = marked_text[:m_flex.start()] + f"<mark>{matched_span}</mark>" + marked_text[m_flex.start() + len(matched_span):]
                has_any_mark = True
                chosen_color = color_hex
                if isinstance(item, dict):
                    item["consumed"] = True
                continue

        # 4. Coincidencia por anclas para frases largas con estricto control de longitud
        words = clean_hl.split()
        matched_full = False
        if len(words) >= 4:
            escaped_words = [re.escape(w) for w in words]
            m_start = None
            for prefix_len in (4, 3, 2):
                p_pat = r"\s+".join(escaped_words[:prefix_len])
                m_start = re.search(p_pat, marked_text, flags=re.IGNORECASE)
                if m_start:
                    break

            if m_start:
                m_end = None
                for suffix_len in (4, 3, 2):
                    s_pat = r"\s+".join(escaped_words[-suffix_len:])
                    m_end = re.search(s_pat, marked_text[m_start.start():], flags=re.IGNORECASE)
                    if m_end:
                        break

                if m_end:
                    abs_end = m_start.start() + m_end.end()
                    matched_span = marked_text[m_start.start():abs_end]
                    # Validar que la longitud de palabras sea coherente (no tragar oraciones siguientes)
                    span_w_count = len(matched_span.split())
                    if abs(span_w_count - len(words)) <= max(2, int(len(words) * 0.20)):
                        marked_text = marked_text[:m_start.start()] + f"<mark>{matched_span}</mark>" + marked_text[abs_end:]
                        has_any_mark = True
                        chosen_color = color_hex
                        matched_full = True
                        if isinstance(item, dict):
                            item["consumed"] = True
                        continue

        # 5. Fallback para fragmentos de subrayado que cruzan párrafos
        # Solo se aplica si la frase cruza literalmente el borde del párrafo (inicia al inicio o termina al final)
        if not matched_full and not (isinstance(item, dict) and item.get("consumed")):
            min_cross_len = max(6, int(len(words) * 0.65))
            if len(words) >= 6:
                matched_span_tuple = None
                for length in range(len(words) - 1, min_cross_len - 1, -1):
                    # Prefijo de palabras al final del párrafo
                    sub_pfx = words[:length]
                    pat_pfx = r"\s+".join(re.escape(w) for w in sub_pfx) + r"\s*$"
                    m = re.search(pat_pfx, marked_text, flags=re.IGNORECASE)
                    if m:
                        matched_span_tuple = (m.start(), m.end())
                        break
                    # Sufijo de palabras al inicio del párrafo
                    sub_sfx = words[-length:]
                    pat_sfx = r"^\s*" + r"\s+".join(re.escape(w) for w in sub_sfx)
                    m = re.search(pat_sfx, marked_text, flags=re.IGNORECASE)
                    if m:
                        matched_span_tuple = (m.start(), m.end())
                        break

                if matched_span_tuple:
                    st_i, end_i = matched_span_tuple
                    sp_text = marked_text[st_i:end_i]
                    if f"<mark>{sp_text}</mark>" not in marked_text:
                        marked_text = marked_text[:st_i] + f"<mark>{sp_text}</mark>" + marked_text[end_i:]
                        has_any_mark = True
                        chosen_color = color_hex
                        matched_full = True
                        continue

    # Limpiar y normalizar marcas
    if has_any_mark:
        marked_text = re.sub(r"</mark>(\s*)<mark>", r"\1", marked_text)
        marked_text = re.sub(r"<mark>\s*</mark>", "", marked_text)
        while "<mark><mark>" in marked_text or "</mark></mark>" in marked_text:
            marked_text = re.sub(r"<mark>(?:\s*<mark>)+", "<mark>", marked_text)
            marked_text = re.sub(r"(?:</mark>\s*)+</mark>", "</mark>", marked_text)

    return marked_text, has_any_mark, chosen_color


def extract_academic_pdf(file_bytes: bytes) -> List[Segment]:
    """Extrae párrafos de PDF preservando número de página real, componentes de portada y secciones."""
    if not PYPDF_AVAILABLE:
        raise UnsupportedFormatError(f"La librería 'pypdf' no está disponible en este entorno: {_pypdf_error}")
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
    except Exception as exc:
        raise CorruptFileError(f"No se pudo leer el archivo PDF: {exc}") from exc

    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise CorruptFileError("El PDF está protegido con contraseña.") from exc

    # Detección de encabezados repetidos entre páginas
    first_lines = []
    for p in reader.pages[:min(6, len(reader.pages))]:
        raw = p.extract_text() or ""
        ls = [l.strip() for l in raw.split("\n") if l.strip()]
        if ls:
            first_lines.append(ls[0])
            if len(ls) > 1:
                first_lines.append(ls[1])

    repeated_headers = {l for l in first_lines if first_lines.count(l) >= 2 or re.match(r"^(?:page\s+\d+|survey\s+on\s+|ieee\s+access|\d+$)", l, re.I)}

    segments: List[Segment] = []
    seg_id = 0
    current_section = "Introducción"
    current_subsection = ""
    section_paragraph_counter = 0

    for page_idx, page in enumerate(reader.pages, start=1):
        # Detección de texto subrayado o resaltado en la página
        page_highlights = _extract_pdf_page_highlights(page)

        try:
            raw_text = page.extract_text() or ""
        except Exception:
            raw_text = ""

        raw_text = _clean_text(raw_text)
        if not raw_text:
            continue

        lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
        split_lines = []
        for l in lines:
            if page_idx == 1:
                m_meta = re.search(r'^(.*?all rights reserved\.?)\s*(?=[A-Z])', l, re.I)
                if m_meta:
                    split_lines.append(m_meta.group(1).strip())
                    split_lines.append(l[m_meta.end():].strip())
                    continue
            split_lines.append(l)

        filtered_lines = []
        for l in split_lines:
            # 1. Limpiar sufijos de encabezado/autores pegados al final de la línea
            l = re.sub(r'\s+[A-Z]\.\s+[A-Za-z\-]+(?:\s+(?:et\s+al\.?|y\s+(?:otros|cols\.?|\bals?\b)))\s*$', '', l)
            l = re.sub(r'\s+Transportation\s+Research\s+Part\s+[A-Z].*$', '', l, flags=re.I)
            l = re.sub(r'\s+I\.\s*Abdulrashid\s+(?:et\s+al\.?|y\s+(?:otros|cols\.?))\s*$', '', l, flags=re.I)

            # 2. Despegar números de página pegados al inicio de línea en páginas > 1
            if page_idx > 1:
                l = re.sub(r'^(?:' + str(page_idx) + r')(?=[a-zA-Z])', '', l)
                l = re.sub(r'^\d{1,3}(?=[a-z])', '', l)

            l = l.strip()
            if not l:
                continue
            if _is_running_header_or_footer(l) or l in repeated_headers or re.match(r"^(?:page\s+\d+(?:\s+of\s+\d+)?|\d+|\d+\s*/\s*\d+)$", l, re.I):
                continue
            filtered_lines.append(l)

        if not filtered_lines:
            continue

        if page_idx == 1:
            raw_tuples = _decompose_page1_academic(filtered_lines)
        else:
            raw_tuples = _parse_body_page_blocks(filtered_lines, current_section)

        for elem_type, sec_hint, text_val in raw_tuples:
            text_val = re.sub(r"[ \t]+", " ", text_val).strip()
            if not text_val or len(text_val) < 2:
                continue

            # Inyectar marcas inline <mark>...</mark> únicamente alrededor de las frases subrayadas
            marked_text_val, is_hl, hl_col = _inject_inline_marks(text_val, page_highlights)

            if elem_type == "heading":
                current_section = sec_hint or text_val
                current_subsection = ""
                section_paragraph_counter = 0
                para_num = 1
            elif elem_type == "metadata":
                current_section = "Encabezado"
                para_num = 1
            elif elem_type == "title":
                current_section = "Título"
                para_num = 1
            elif elem_type == "authors":
                current_section = "Autores"
                para_num = 1
            elif elem_type == "affiliations":
                current_section = "Afiliaciones"
                para_num = 1
            elif elem_type == "abstract":
                current_section = "Resumen"
                para_num = 1
            elif elem_type == "keywords":
                current_section = "Palabras Clave"
                para_num = 1
            elif elem_type in ("table", "caption"):
                section_paragraph_counter += 1
                para_num = section_paragraph_counter
            else:
                section_paragraph_counter += 1
                para_num = section_paragraph_counter

            seg = Segment(
                id=seg_id,
                original=marked_text_val,
                section=current_section,
                subsection=current_subsection,
                page=page_idx,
                paragraph_num=para_num,
                element_type=elem_type,
                is_marked=bool(is_hl or "<mark>" in marked_text_val or "<u>" in marked_text_val),
                color=hl_col,
            )
            segments.append(seg)
            seg_id += 1

    if not segments:
        raise CorruptFileError("No se pudo extraer texto del PDF (posible PDF escaneado sin capa OCR).")
    return segments


def extract_academic_docx(file_bytes: bytes) -> List[Segment]:
    """Extrae párrafos de DOCX detectando estilos de encabezado y estimando páginas."""
    if not DOCX_AVAILABLE:
        raise UnsupportedFormatError(f"La librería 'python-docx' no está disponible en este entorno: {_docx_error}")
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
    except Exception as exc:
        raise CorruptFileError(f"No se pudo leer el archivo DOCX: {exc}") from exc

    segments: List[Segment] = []
    seg_id = 0
    current_section = "Introducción"
    current_subsection = ""
    current_page = 1
    words_in_current_page = 0
    section_paragraph_counter = 0

    for idx, p in enumerate(doc.paragraphs):
        text = _clean_text(p.text)
        if not text:
            continue

        # Detectar saltos de página explícitos en DOCX
        has_page_break = any(
            run._r.xpath("./w:br[@w:type='page']") for run in p.runs
        ) if hasattr(p, "runs") else False

        if has_page_break:
            current_page += 1
            words_in_current_page = 0

        # Estimar avance de página por volumen de palabras (estándar académico: ~450 palabras/página)
        word_count = len(text.split())
        words_in_current_page += word_count
        if words_in_current_page > 480 and not has_page_break:
            current_page += 1
            words_in_current_page = word_count

        style_name = (p.style.name if p.style else "").lower()

        elem_type, new_sec, new_subsec = _detect_element_type(
            text, current_section, current_page, is_first=(idx == 0)
        )

        if "heading 1" in style_name or "title" in style_name:
            new_sec = text
            elem_type = "heading"
        elif "heading 2" in style_name or "heading 3" in style_name:
            new_subsec = text
            elem_type = "heading"

        if new_sec:
            current_section = new_sec
            current_subsection = ""
            section_paragraph_counter = 0
        if new_subsec:
            current_subsection = new_subsec

        if elem_type != "heading":
            section_paragraph_counter += 1
            para_num = section_paragraph_counter
        else:
            para_num = 1

        # Detección de texto resaltado o subrayado en runs de DOCX a nivel de frase
        has_any_mark = False
        run_highlight_color = "#FEF08A"
        marked_doc_text = text
        if hasattr(p, "runs") and p.runs:
            run_parts = []
            for r in p.runs:
                r_text = r.text
                if not r_text:
                    continue
                is_r_marked = (r.font.highlight_color is not None or (r.font.underline is not None and r.font.underline is not False))
                if is_r_marked:
                    has_any_mark = True
                    run_parts.append(f"<mark>{r_text}</mark>")
                    if r.font.highlight_color:
                        hl_name = str(r.font.highlight_color).upper()
                        if "YELLOW" in hl_name:
                            run_highlight_color = "#FEF08A"
                        elif "GREEN" in hl_name:
                            run_highlight_color = "#BBF7D0"
                        elif "CYAN" in hl_name or "TURQUOISE" in hl_name:
                            run_highlight_color = "#BAE6FD"
                        elif "PINK" in hl_name or "MAGENTA" in hl_name:
                            run_highlight_color = "#FBCFE8"
                else:
                    run_parts.append(r_text)
            if has_any_mark:
                marked_doc_text = "".join(run_parts)
                marked_doc_text = re.sub(r"</mark>(\s*)<mark>", r"\1", marked_doc_text)
                marked_doc_text = re.sub(r"<mark>\s*</mark>", "", marked_doc_text)

        seg = Segment(
            id=seg_id,
            original=marked_doc_text,
            section=current_section,
            subsection=current_subsection,
            page=current_page,
            paragraph_num=para_num,
            element_type=elem_type,
            is_marked=has_any_mark,
            color=run_highlight_color,
        )
        segments.append(seg)
        seg_id += 1

    if not segments:
        raise CorruptFileError("El documento DOCX no contiene texto extraíble.")
    return segments


def extract_academic_txt(file_bytes: bytes) -> List[Segment]:
    """Extrae párrafos de TXT identificando marcadores de página o encabezados."""
    try:
        text = file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = file_bytes.decode("latin-1", errors="ignore")

    text = _clean_text(text)
    raw_blocks = [p.strip() for p in text.split("\n\n") if p.strip()]

    segments: List[Segment] = []
    seg_id = 0
    current_section = "Introducción"
    current_subsection = ""
    current_page = 1
    words_in_current_page = 0
    section_paragraph_counter = 0

    page_marker_regex = re.compile(r"^(?:---\s*)?(?:Page|Página|Pag\.)\s*(\d+)(?:\s*---)?$", re.IGNORECASE)

    for idx, block in enumerate(raw_blocks):
        # Verificar si es un marcador explícito de página (ej. "--- Page 2 ---")
        p_match = page_marker_regex.match(block)
        if p_match:
            current_page = int(p_match.group(1))
            words_in_current_page = 0
            continue

        words = len(block.split())
        words_in_current_page += words
        if words_in_current_page > 450:
            current_page += 1
            words_in_current_page = words

        elem_type, new_sec, new_subsec = _detect_element_type(
            block, current_section, current_page, is_first=(idx == 0)
        )

        if new_sec:
            current_section = new_sec
            current_subsection = ""
            section_paragraph_counter = 0
        if new_subsec:
            current_subsection = new_subsec

        if elem_type != "heading":
            section_paragraph_counter += 1
            para_num = section_paragraph_counter
        else:
            para_num = 1

        seg = Segment(
            id=seg_id,
            original=block,
            section=current_section,
            subsection=current_subsection,
            page=current_page,
            paragraph_num=para_num,
            element_type=elem_type,
        )
        segments.append(seg)
        seg_id += 1

    if not segments:
        raise CorruptFileError("El archivo de texto está vacío.")
    return segments


def extract_academic_segments(filename: str, file_bytes: bytes) -> List[Segment]:
    """Punto de entrada universal para extracción estructurada académica."""
    ext = "." + filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if ext == ".txt":
        return extract_academic_txt(file_bytes)
    if ext == ".docx":
        return extract_academic_docx(file_bytes)
    if ext == ".pdf":
        return extract_academic_pdf(file_bytes)
    raise UnsupportedFormatError(f"Formato no soportado: {ext or 'desconocido'}")


# --------------------------------------------------------------------------
# Compatibilidad hacia atrás con la función simple original
# --------------------------------------------------------------------------

def extract_text_segments(filename: str, file_bytes: bytes) -> List[str]:
    """Función de compatibilidad para código anterior."""
    segs = extract_academic_segments(filename, file_bytes)
    return [s.original for s in segs]


# --------------------------------------------------------------------------
# Exportación Normal y Enriquecida (.txt, .docx, .pdf)
# --------------------------------------------------------------------------

def export_txt(segments: List[Segment], enriched: bool = False) -> bytes:
    """Exporta los segmentos a TXT. Si enriched=True, añade fichas de procedencia en párrafos marcados."""
    lines = []
    for s in segments:
        text = s.translated or s.original
        clean_text = re.sub(r"</?(?:mark|u)>", "", text)
        if s.is_marked:
            lines.append(f"★ [SUBRAYADO EN EL ORIGINAL / CITACIÓN]")
            lines.append(clean_text)
            if enriched:
                lines.append(f"   ↳ [Procedencia: {s.provenance_label}]")
        else:
            lines.append(clean_text)
    return "\n\n".join(lines).encode("utf-8")


def _add_docx_tagged_runs(p, text: str, font_size_pt: float = 9.5, is_globally_marked: bool = False, italic: bool = False, bold: bool = False):
    """Agrega texto a un párrafo DOCX aplicando SUBRAYADO (underline) y resaltado a las frases dentro de <mark>...</mark> o <u>...</u>."""
    if not text:
        return
    if "<mark>" in text or "<u>" in text:
        parts = re.split(r"(</?(?:mark|u)>)", text)
        in_mark = False
        for part in parts:
            if part in ("<mark>", "<u>"):
                in_mark = True
            elif part in ("</mark>", "</u>"):
                in_mark = False
            elif part:
                run = p.add_run(part)
                run.font.size = Pt(font_size_pt)
                if italic:
                    run.italic = True
                if bold:
                    run.bold = True
                if in_mark:
                    run.font.underline = True
                    if WD_COLOR_INDEX and hasattr(WD_COLOR_INDEX, "YELLOW"):
                        run.font.highlight_color = WD_COLOR_INDEX.YELLOW
    else:
        run = p.add_run(text)
        run.font.size = Pt(font_size_pt)
        if italic:
            run.italic = True
        if bold:
            run.bold = True


def export_docx(segments: List[Segment], enriched: bool = False) -> bytes:
    """
    Exporta a .DOCX con estructura académica fiel al formato de paper científico:
      - Título centrado y destacado
      - Abstract en 1 columna con márgenes indentados y cursiva
      - Cuerpo del paper en formato de DOS COLUMNAS (estándar IEEE/revista académica)
      - Encabezados de sección (H1, H2)
      - Referencias bibliográficas formateadas
      - Resaltado y nota de procedencia si enriched=True
    """
    if not DOCX_AVAILABLE:
        raise UnsupportedFormatError(f"No se puede generar .docx (librería 'python-docx' no disponible): {_docx_error}")

    try:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.enum.section import WD_SECTION
    except ImportError:
        OxmlElement = qn = WD_SECTION = None

    doc = DocxDocument()

    # 1. Separar Título, Autores, Afiliaciones, Abstract y Cuerpo
    metadata_segs = [s for s in segments if s.element_type == "metadata"]
    title_segs = [s for s in segments if s.element_type == "title"]
    authors_segs = [s for s in segments if s.element_type == "authors"]
    affiliations_segs = [s for s in segments if s.element_type == "affiliations"]
    abstract_segs = [s for s in segments if s.element_type == "abstract" or s.section.lower() in ["abstract", "resumen"]]
    keywords_segs = [s for s in segments if s.element_type == "keywords"]
    
    consumed_ids = {s.id for s in (metadata_segs + title_segs + authors_segs + affiliations_segs + abstract_segs + keywords_segs)}
    body_segs = [s for s in segments if s.id not in consumed_ids]

    # Sección 1 (1 columna): Metadatos, Título, Autores, Afiliaciones y Resumen / Abstract
    if metadata_segs:
        meta_p = doc.add_paragraph()
        meta_p.paragraph_format.space_after = Pt(4)
        run_meta = meta_p.add_run(" · ".join((s.translated or s.original) for s in metadata_segs))
        run_meta.font.size = Pt(8)
        run_meta.font.color.rgb = RGBColor(100, 116, 139)

    if title_segs:
        for ts in title_segs:
            p = doc.add_heading(ts.translated or ts.original, level=0)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(10)
    elif segments and not body_segs:
        first = segments[0]
        p = doc.add_heading(first.translated or first.original, level=0)
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(10)

    # Autores
    if authors_segs:
        auth_p = doc.add_paragraph()
        auth_p.paragraph_format.alignment = 1  # Centrado
        auth_p.paragraph_format.space_after = Pt(4)
        run_auth = auth_p.add_run(" · ".join((s.translated or s.original) for s in authors_segs))
        run_auth.bold = True
        run_auth.font.size = Pt(10)

    # Afiliaciones
    if affiliations_segs:
        affil_p = doc.add_paragraph()
        affil_p.paragraph_format.alignment = 1  # Centrado
        affil_p.paragraph_format.space_after = Pt(12)
        run_affil = affil_p.add_run(" — ".join((s.translated or s.original) for s in affiliations_segs))
        run_affil.italic = True
        run_affil.font.size = Pt(8.5)
        run_affil.font.color.rgb = RGBColor(100, 116, 139)

    # Abstract / Resumen
    if abstract_segs:
        for ab in abstract_segs:
            text = ab.translated or ab.original
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.45)
            p.paragraph_format.right_indent = Inches(0.45)
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(8)
            run_bold = p.add_run("RESUMEN — ")
            run_bold.bold = True
            run_bold.font.size = Pt(9.5)
            clean_abs = re.sub(r"^(?:A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|abstract|R\s*E\s*S\s*U\s*M\s*E\s*N|resumen)\s*[\:\—\-\.]*\s*", "", text, flags=re.I)
            _add_docx_tagged_runs(p, clean_abs, font_size_pt=9.5, italic=True)

    # Palabras clave
    if keywords_segs:
        for kw in keywords_segs:
            text = kw.translated or kw.original
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.45)
            p.paragraph_format.right_indent = Inches(0.45)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(14)
            run_kw_bold = p.add_run("PALABRAS CLAVE — ")
            run_kw_bold.bold = True
            run_kw_bold.font.size = Pt(8.8)
            clean_kw = re.sub(r"^(?:K\s*E\s*Y\s*W\s*O\s*R\s*D\s*S|keywords|index\s+terms|palabras\s+clave|key\s+words|t[eé]rminos\s+de\s+[ií]ndice)\s*[\:\—\-\.]*\s*", "", text, flags=re.I)
            _add_docx_tagged_runs(p, clean_kw, font_size_pt=8.8, italic=True)

    # Sección 2 (2 columnas para el cuerpo del artículo estilo IEEE)
    body_section = doc.add_section(WD_SECTION.CONTINUOUS)
    body_sectPr = body_section._sectPr
    cols = body_sectPr.xpath('./w:cols')
    if cols:
        cols[0].set(qn('w:num'), '2')
        cols[0].set(qn('w:space'), '720')  # 0.5 pulgada entre columnas
    else:
        col_elem = OxmlElement('w:cols')
        col_elem.set(qn('w:num'), '2')
        col_elem.set(qn('w:space'), '720')
        body_sectPr.append(col_elem)

    # Renderizar segmentos en las dos columnas
    for s in body_segs:
        text = s.translated or s.original
        elem_type = s.element_type

        if elem_type == "heading":
            level = 2 if s.subsection else 1
            p = doc.add_heading(text, level=level)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
        elif elem_type == "caption":
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            _add_docx_tagged_runs(p, text, font_size_pt=8.5, italic=True)
        elif elem_type == "table":
            t_lines = [l.strip() for l in text.split("\n") if l.strip()]
            if t_lines:
                cap_p = doc.add_paragraph()
                cap_p.paragraph_format.space_before = Pt(8)
                cap_p.paragraph_format.space_after = Pt(2)
                cap_run = cap_p.add_run(t_lines[0])
                cap_run.bold = True
                cap_run.font.size = Pt(8.5)
                for tl in t_lines[1:]:
                    tp = doc.add_paragraph()
                    tp.paragraph_format.space_after = Pt(1)
                    trun = tp.add_run(tl)
                    trun.font.name = "Consolas"
                    trun.font.size = Pt(7.5)
        elif elem_type == "reference":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.first_line_indent = Inches(-0.3)
            p.paragraph_format.space_after = Pt(3)
            _add_docx_tagged_runs(p, text, font_size_pt=8, is_globally_marked=False)
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15
            _add_docx_tagged_runs(p, text, font_size_pt=9.5, is_globally_marked=False)

        # Si está marcado y enriquecido, nota de procedencia
        if enriched and s.is_marked and elem_type != "heading":
            prov_p = doc.add_paragraph()
            prov_p.paragraph_format.left_indent = Inches(0.2)
            prov_p.paragraph_format.space_after = Pt(6)
            prov_run = prov_p.add_run(f"📌 Procedencia: {s.provenance_label}")
            prov_run.italic = True
            prov_run.font.size = Pt(8)
            prov_run.font.color.rgb = RGBColor(79, 70, 229)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _clean_pdf_text(text: str) -> str:
    """Sanitiza caracteres tipográficos especiales, matemáticas y ligaduras para compatibilidad total con fuentes estándar de ReportLab."""
    if not text:
        return ""
    text = repair_academic_symbols_and_ligatures(text)
    text = clean_math_display(text)

    char_map = {
        "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'",
        "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u201f": '"',
        "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": " - ", "\u2014": " -- ", "\u2015": " -- ",
        "\u2026": "...", "\xa0": " ", "\u202f": " ",
        "\u2264": "<=", "\u2265": ">=", "\u2260": "!=", "\u2248": "~", "\u00d7": "x", "\u00f7": "/",
        "\u2208": " in ", "\u2209": " not in ", "\u2211": "sum ", "\u220f": "prod ",
        "\u221e": "inf", "\u2192": "->", "\u2190": "<-", "\u2194": "<->",
        "\u27e8": "<", "\u27e9": ">", "\u3008": "<", "\u3009": ">",
        "\u22c5": "*", "\u22c6": "*", "\u223c": "~",
        "\u22a8": "|=", "\u25a0": "-", "\u25a1": "-", "\ufffd": "",
        # Letras griegas a representación legible estándar
        "α": "alpha", "β": "beta", "γ": "gamma", "δ": "delta", "ε": "epsilon", "ζ": "zeta", "η": "eta",
        "θ": "theta", "ι": "iota", "κ": "kappa", "λ": "lambda", "μ": "mu", "ν": "nu", "ξ": "xi",
        "π": "pi", "ρ": "rho", "σ": "sigma", "τ": "tau", "υ": "upsilon", "φ": "phi", "χ": "chi",
        "ψ": "psi", "ω": "omega",
        "Α": "Alpha", "Β": "Beta", "Γ": "Gamma", "Δ": "Delta", "Ε": "Epsilon", "Ζ": "Zeta", "Η": "Eta",
        "Θ": "Theta", "Ι": "Iota", "Κ": "Kappa", "Λ": "Lambda", "Μ": "Mu", "Ν": "Nu", "Ξ": "Xi",
        "Π": "Pi", "Ρ": "Rho", "Σ": "Sigma", "Τ": "Tau", "Υ": "Upsilon", "Φ": "Phi", "Χ": "Chi",
        "Ψ": "Psi", "Ω": "Omega"
    }
    for k, v in char_map.items():
        text = text.replace(k, v)

    # Filtrar caracteres con ord > 255 que provocarían errores en ReportLab Helvetica
    sanitized = []
    for ch in text:
        if ord(ch) <= 255:
            sanitized.append(ch)
        else:
            norm = unicodedata.normalize('NFKD', ch)
            norm_latin = "".join(c for c in norm if ord(c) <= 255)
            sanitized.append(norm_latin if norm_latin else " ")
    return "".join(sanitized).strip()


def _draw_academic_line(c, text: str, x: float, y: float, font_name: str, font_size: float, target_width: float, is_last_line: bool = False):
    """Dibuja una línea de texto con justificación automática si no es la última línea del párrafo."""
    words = text.split()
    if is_last_line or len(words) <= 1:
        c.drawString(x, y, text)
        return

    try:
        text_width = c.stringWidth(text, font_name, font_size)
        extra_space = target_width - text_width
        # Si el espacio extra es moderado (menos del 22% del ancho), justificar usando setWordSpace
        if 0 < extra_space < (target_width * 0.22):
            word_space = extra_space / (len(words) - 1)
            t = c.beginText(x, y)
            t.setFont(font_name, font_size)
            t.setWordSpace(word_space)
            t.textLine(text)
            c.drawText(t)
        else:
            c.drawString(x, y, text)
    except Exception:
        c.drawString(x, y, text)


def _split_tagged_words(text: str) -> List[Tuple[str, bool]]:
    """Descompone texto con etiquetas <mark>...</mark> o <u>...</u> en palabras etiquetadas (palabra, es_resaltada)."""
    tokens = re.split(r"(</?(?:mark|u)>)", text)
    tagged_words = []
    is_m = False
    for tok in tokens:
        if tok in ("<mark>", "<u>"):
            is_m = True
        elif tok in ("</mark>", "</u>"):
            is_m = False
        else:
            for w in tok.split():
                clean_w = re.sub(r"</?(?:mark|u)>", "", w)
                if clean_w:
                    tagged_words.append((clean_w, is_m))
    return tagged_words


def _wrap_tagged_lines(c, tagged_words: List[Tuple[str, bool]], font_name: str, font_size: float, target_width: float) -> List[List[Tuple[str, bool]]]:
    """Ajusta palabras etiquetadas en líneas que se adaptan al ancho de columna objetivo."""
    lines = []
    curr_line = []
    curr_w = 0.0
    space_w = c.stringWidth(" ", font_name, font_size)

    for word, m in tagged_words:
        w_w = c.stringWidth(word, font_name, font_size)
        needed = w_w if not curr_line else (space_w + w_w)
        if curr_line and (curr_w + needed > target_width):
            lines.append(curr_line)
            curr_line = [(word, m)]
            curr_w = w_w
        else:
            curr_line.append((word, m))
            curr_w += needed
    if curr_line:
        lines.append(curr_line)
    return lines


def _draw_tagged_line(c, line: List[Tuple[str, bool]], x: float, y: float, font_name: str, font_size: float, target_width: float, is_last_line: bool = False, text_color: Tuple[float, float, float] = (0.08, 0.08, 0.10)):
    """Dibuja una línea con justificación tipográfica, fondo suave y SUBRAYADO nítido ÚNICAMENTE en las palabras marcadas."""
    if not line:
        return
    space_w = c.stringWidth(" ", font_name, font_size)
    num_words = len(line)

    if not is_last_line and num_words > 1:
        total_word_w = sum(c.stringWidth(w, font_name, font_size) for w, _ in line)
        standard_total = total_word_w + (num_words - 1) * space_w
        extra = target_width - standard_total
        if 0 < extra < (target_width * 0.22):
            actual_space_w = space_w + (extra / (num_words - 1))
        else:
            actual_space_w = space_w
    else:
        actual_space_w = space_w

    # Paso 1: Detectar tramos marcados
    curr_x = x
    hl_spans = []
    span_start = None
    span_end = None

    for word, is_m in line:
        w_w = c.stringWidth(word, font_name, font_size)
        word_start = curr_x
        word_end = curr_x + w_w

        if is_m:
            if span_start is None:
                span_start = word_start
            span_end = word_end
        else:
            if span_start is not None:
                hl_spans.append((span_start, span_end - span_start))
                span_start = None
                span_end = None

        curr_x += w_w + actual_space_w

    if span_start is not None:
        hl_spans.append((span_start, span_end - span_start))

    # Paso 2: Dibujar fondo de resaltado suave SOLO en los tramos marcados
    for hx, hw in hl_spans:
        c.saveState()
        c.setFillColorRGB(1.0, 0.95, 0.60)  # Amarillo fluorescente suave de marcador
        c.rect(hx - 1, y - 2, hw + 2, font_size + 3.0, fill=1, stroke=0)
        c.restoreState()

    # Paso 3: Dibujar el texto
    curr_x = x
    c.setFont(font_name, font_size)
    c.setFillColorRGB(*text_color)
    for word, _ in line:
        c.drawString(curr_x, y, word)
        curr_x += c.stringWidth(word, font_name, font_size) + actual_space_w

    # Paso 4: Dibujar línea de SUBRAYADO tipográfico nítida y visible directamente bajo la línea de base
    for hx, hw in hl_spans:
        c.saveState()
        c.setStrokeColorRGB(0.70, 0.38, 0.05)  # Color ámbar académico visible
        c.setLineWidth(1.15)
        c.line(hx - 0.5, y - 1.8, hx + hw + 0.5, y - 1.8)
        c.restoreState()


def export_pdf(segments: List[Segment], enriched: bool = False) -> bytes:
    """
    Exporta a .PDF en formato académico de DOS COLUMNAS balanceadas (estilo IEEE / revista científica):
      - Encabezado institucional en la parte superior
      - Título centrado a ancho completo con tipografía destacada
      - Autores y afiliaciones bien separados
      - Abstract / Resumen en caja sombreada a ancho completo
      - Línea divisoria decorativa
      - Flujo continuo línea por línea en DOS COLUMNAS (evita saltos vacíos y desalineaciones)
      - Encabezados vinculados a su texto (evita encabezados huérfanos)
      - Citas y fórmulas matemáticas higienizadas y legibles
      - Párrafos marcados con barra ámbar lateral, fondo suave y trazabilidad de origen
      - Encabezados y números de página continuos en cada hoja
    """
    if not REPORTLAB_AVAILABLE:
        return export_txt(segments, enriched=enriched)

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    margin = 38.0
    col_gap = 18.0
    col_width = (width - 2 * margin - col_gap) / 2
    bottom_margin = 42.0

    # Separar Componentes Estructurales de Portada
    metadata_segs = [s for s in segments if s.element_type == "metadata"]
    title_segs = [s for s in segments if s.element_type == "title"]
    authors_segs = [s for s in segments if s.element_type == "authors"]
    affiliations_segs = [s for s in segments if s.element_type == "affiliations"]
    abstract_segs = [s for s in segments if s.element_type == "abstract" or s.section.lower() in ["abstract", "resumen"]]
    keywords_segs = [s for s in segments if s.element_type == "keywords"]

    consumed_ids = {s.id for s in (metadata_segs + title_segs + authors_segs + affiliations_segs + abstract_segs + keywords_segs)}
    body_segs = [s for s in segments if s.id not in consumed_ids]

    metadata_text = _clean_pdf_text(" · ".join((s.translated or s.original) for s in metadata_segs)) if metadata_segs else ""
    title_text = _clean_pdf_text(title_segs[0].translated or title_segs[0].original) if title_segs else ""
    if not title_text and segments:
        title_text = _clean_pdf_text(segments[0].translated or segments[0].original)
        if segments[0] in body_segs:
            body_segs.remove(segments[0])

    authors_text = _clean_pdf_text(" · ".join((s.translated or s.original) for s in authors_segs)) if authors_segs else ""
    affiliations_text = _clean_pdf_text(" — ".join((s.translated or s.original) for s in affiliations_segs)) if affiliations_segs else ""
    abstract_text = _clean_pdf_text(" ".join((s.translated or s.original) for s in abstract_segs)) if abstract_segs else ""
    keywords_text = _clean_pdf_text(" ".join((s.translated or s.original) for s in keywords_segs)) if keywords_segs else ""

    # --------------------------------------------------------------------------
    # PÁGINA 1: Encabezado superior, Título y Abstract
    # --------------------------------------------------------------------------
    # 1. Banner superior
    c.setFont("Helvetica-Bold", 7.0)
    c.setFillColorRGB(0.35, 0.38, 0.45)
    banner_left = metadata_text[:55] if metadata_text else "TRADUCCIÓN ACADÉMICA CON TRAZABILIDAD DE ORIGEN"
    c.drawString(margin, height - 26, banner_left.upper())
    c.drawRightString(width - margin, height - 26, "REVISTA CIENTÍFICA · FORMATO IEEE")
    c.setStrokeColorRGB(0.78, 0.82, 0.88)
    c.setLineWidth(0.6)
    c.line(margin, height - 30, width - margin, height - 30)

    # 2. Título centrado
    y = height - 50
    c.setFont("Helvetica-Bold", 14.0)
    c.setFillColorRGB(0.08, 0.10, 0.18)
    title_lines = simpleSplit(title_text, "Helvetica-Bold", 14.0, width - 2 * margin - 20)
    for tl in title_lines:
        c.drawCentredString(width / 2, y, tl)
        y -= 17

    # 3. Autores
    if authors_text:
        y -= 3
        c.setFont("Helvetica-Bold", 9.0)
        c.setFillColorRGB(0.12, 0.16, 0.24)
        author_lines = simpleSplit(authors_text, "Helvetica-Bold", 9.0, width - 2 * margin - 20)
        for al in author_lines:
            c.drawCentredString(width / 2, y, al)
            y -= 12

    # 4. Afiliaciones
    if affiliations_text:
        y -= 1
        c.setFont("Helvetica-Oblique", 7.5)
        c.setFillColorRGB(0.40, 0.44, 0.52)
        affil_lines = simpleSplit(affiliations_text, "Helvetica-Oblique", 7.5, width - 2 * margin - 30)
        for afl in affil_lines:
            c.drawCentredString(width / 2, y, afl)
            y -= 10.5

    # 5. Caja de Resumen / Abstract y Palabras Clave
    if abstract_text or keywords_text:
        y -= 4
        clean_abs = re.sub(r"^(?:A\s*B\s*S\s*T\s*R\s*A\s*C\s*T|abstract|R\s*E\s*S\s*U\s*M\s*E\s*N|resumen)\s*[\:\—\-\.]*\s*", "", abstract_text, flags=re.I).strip()
        abs_width = width - 2 * margin - 22

        has_inline_abs = ("<mark>" in clean_abs or "<u>" in clean_abs)
        abs_tagged_lines = []
        abs_plain_lines = []

        if clean_abs:
            if has_inline_abs:
                abs_tagged_content = "RESUMEN — " + clean_abs
                t_words = _split_tagged_words(abs_tagged_content)
                abs_tagged_lines = _wrap_tagged_lines(c, t_words, "Helvetica-Oblique", 8.2, abs_width)
                num_abs_lines = len(abs_tagged_lines)
            else:
                clean_plain_abs = re.sub(r"</?(?:mark|u)>", "", clean_abs)
                abs_full = "RESUMEN — " + clean_plain_abs
                c.setFont("Helvetica-Oblique", 8.2)
                abs_plain_lines = simpleSplit(abs_full, "Helvetica-Oblique", 8.2, abs_width)
                num_abs_lines = len(abs_plain_lines)
        else:
            num_abs_lines = 0

        kw_lines = []
        if keywords_text:
            clean_kw = re.sub(r"^(?:K\s*E\s*Y\s*W\s*O\s*R\s*D\s*S|keywords|index\s+terms|palabras\s+clave|key\s+words|t[eé]rminos\s+de\s+[ií]ndice)\s*[\:\—\-\.]*\s*", "", keywords_text, flags=re.I).strip()
            clean_kw = re.sub(r"</?(?:mark|u)>", "", clean_kw)
            if clean_kw:
                kw_full = "PALABRAS CLAVE — " + clean_kw
                kw_lines = simpleSplit(kw_full, "Helvetica-BoldOblique", 7.8, abs_width)

        box_padding = 8
        abs_box_height = (num_abs_lines * 11.0) + (len(kw_lines) * 10.2 + 4 if kw_lines else 0) + (2 * box_padding)

        abs_is_marked = any(ab.is_marked for ab in abstract_segs)
        c.saveState()
        if abs_is_marked:
            c.setFillColorRGB(1.0, 0.98, 0.90)
            c.setStrokeColorRGB(0.92, 0.70, 0.12)
        else:
            c.setFillColorRGB(0.96, 0.97, 0.99)
            c.setStrokeColorRGB(0.80, 0.84, 0.90)
        c.roundRect(margin, y - abs_box_height, width - 2 * margin, abs_box_height, 4, fill=1, stroke=1)
        c.restoreState()

        y_abs = y - box_padding - 8
        c.setFillColorRGB(0.12, 0.14, 0.20)

        if has_inline_abs and abs_tagged_lines:
            for l_idx, tl in enumerate(abs_tagged_lines):
                is_last = (l_idx == len(abs_tagged_lines) - 1)
                _draw_tagged_line(c, tl, margin + 11, y_abs, "Helvetica-Oblique", 8.2, abs_width, is_last_line=is_last, text_color=(0.12, 0.14, 0.20))
                y_abs -= 11.0
        elif abs_plain_lines:
            c.setFont("Helvetica-Oblique", 8.2)
            for al in abs_plain_lines:
                c.drawString(margin + 11, y_abs, al)
                y_abs -= 11.0

        if kw_lines:
            y_abs -= 3
            c.setFont("Helvetica-BoldOblique", 7.8)
            c.setFillColorRGB(0.20, 0.25, 0.35)
            for kl in kw_lines:
                c.drawString(margin + 11, y_abs, kl)
                y_abs -= 10.2

        y -= abs_box_height + 12
    else:
        y -= 8

    # Línea divisoria antes de iniciar las 2 columnas
    c.setStrokeColorRGB(0.82, 0.85, 0.90)
    c.setLineWidth(0.5)
    c.line(margin, y, width - margin, y)
    y -= 14

    # --------------------------------------------------------------------------
    # FLUJO CONTINUO LÍNEA POR LÍNEA EN DOS COLUMNAS
    # --------------------------------------------------------------------------
    col_top_y_p1 = y
    page_num = 1
    current_col = 0  # 0: izquierda, 1: derecha
    curr_y = col_top_y_p1

    def get_col_x(col_idx: int) -> float:
        return margin if col_idx == 0 else (margin + col_width + col_gap)

    def draw_running_header_footer(pg: int):
        c.setFont("Helvetica", 7.5)
        c.setFillColorRGB(0.42, 0.46, 0.54)
        if pg > 1:
            short_t = title_text[:60] + ("..." if len(title_text) > 60 else "")
            c.drawString(margin, height - 24, short_t.upper())
            c.drawRightString(width - margin, height - 24, f"Pág. {pg}")
            c.setStrokeColorRGB(0.82, 0.85, 0.90)
            c.setLineWidth(0.5)
            c.line(margin, height - 28, width - margin, height - 28)
        c.drawCentredString(width / 2, 22, f"— Página {pg} —")

    def next_column_or_page():
        nonlocal current_col, curr_y, page_num
        if current_col == 0:
            current_col = 1
            curr_y = col_top_y_p1 if page_num == 1 else (height - margin - 22)
        else:
            c.showPage()
            page_num += 1
            draw_running_header_footer(page_num)
            current_col = 0
            curr_y = height - margin - 22

    draw_running_header_footer(1)

    for s in body_segs:
        raw_text = s.translated or s.original
        text = _clean_pdf_text(raw_text)
        if not text:
            continue

        elem_type = s.element_type
        is_marked = s.is_marked
        show_provenance = (enriched and s.is_marked)

        # Configuración tipográfica según jerarquía del paper
        if elem_type == "heading":
            is_sub = bool(s.subsection)
            font_name = "Helvetica-BoldOblique" if is_sub else "Helvetica-Bold"
            font_size = 8.8 if is_sub else 10.0
            leading = 12.0 if is_sub else 13.5
            space_before = 8.0 if is_sub else 11.0
            space_after = 4.0
            text_color = (0.15, 0.20, 0.35)
        elif elem_type == "caption":
            font_name = "Helvetica-Oblique"
            font_size = 7.8
            leading = 10.2
            space_before = 4.0
            space_after = 6.0
            text_color = (0.25, 0.30, 0.40)
        elif elem_type == "table":
            font_name = "Courier"
            font_size = 6.8
            leading = 8.8
            space_before = 5.0
            space_after = 6.0
            text_color = (0.10, 0.12, 0.15)
        elif elem_type == "reference":
            font_name = "Helvetica"
            font_size = 7.5
            leading = 9.8
            space_before = 2.0
            space_after = 3.0
            text_color = (0.18, 0.20, 0.25)
        else:
            font_name = "Helvetica"
            font_size = 8.6
            leading = 11.6
            space_before = 2.0
            space_after = 5.0
            text_color = (0.08, 0.08, 0.10)

        # 1. Si es encabezado, asegurar que no quede huérfano al final de columna
        if elem_type == "heading":
            clean_head = re.sub(r"</?(?:mark|u)>", "", text)
            lines = simpleSplit(clean_head, font_name, font_size, col_width)
            if not lines:
                continue
            needed_heading_space = space_before + (len(lines) * leading) + (3 * 11.6) + space_after
            if curr_y - needed_heading_space < bottom_margin:
                next_column_or_page()

            curr_y -= space_before
            col_x = get_col_x(current_col)
            c.setFont(font_name, font_size)
            c.setFillColorRGB(*text_color)
            for hl in lines:
                c.drawString(col_x, curr_y, hl)
                curr_y -= leading
            curr_y -= space_after
            continue

        # 2. Párrafo normal o referencia con soporte de resaltado inline exacto
        has_inline_mark = ("<mark>" in raw_text or "<u>" in raw_text)
        if has_inline_mark:
            tagged_words = _split_tagged_words(raw_text)
            wrapped_lines = _wrap_tagged_lines(c, tagged_words, font_name, font_size, col_width)
            if not wrapped_lines:
                continue

            curr_y -= space_before
            for line_idx, line in enumerate(wrapped_lines):
                if curr_y - leading < bottom_margin:
                    next_column_or_page()

                col_x = get_col_x(current_col)
                is_last = (line_idx == len(wrapped_lines) - 1)
                _draw_tagged_line(c, line, col_x, curr_y, font_name, font_size, col_width, is_last_line=is_last, text_color=text_color)
                curr_y -= leading
        else:
            clean_plain = re.sub(r"</?(?:mark|u)>", "", text)
            lines = simpleSplit(clean_plain, font_name, font_size, col_width)
            if not lines:
                continue

            curr_y -= space_before
            for line_idx, line in enumerate(lines):
                if curr_y - leading < bottom_margin:
                    next_column_or_page()

                col_x = get_col_x(current_col)
                c.setFont(font_name, font_size)
                c.setFillColorRGB(*text_color)
                is_last = (line_idx == len(lines) - 1)
                _draw_academic_line(c, line, col_x, curr_y, font_name, font_size, col_width, is_last_line=is_last)
                curr_y -= leading

        # Nota de procedencia solo si el párrafo está marcado y es enriquecido
        if show_provenance:
            prov_leading = 9.5
            if curr_y - prov_leading < bottom_margin:
                next_column_or_page()
            col_x = get_col_x(current_col)
            c.setFont("Helvetica-Oblique", 7.2)
            c.setFillColorRGB(0.26, 0.22, 0.75)  # Índigo académico
            c.drawString(col_x, curr_y, f"[📌 Origen: {s.provenance_label}]")
            curr_y -= prov_leading + 2

        curr_y -= space_after

    c.save()
    return buf.getvalue()


def export_segments(segments: List[Segment], target_format: str, enriched: bool = False) -> bytes:
    """Función de exportación unificada con soporte de modo enriquecido."""
    target_format = target_format.lower().lstrip(".")
    if target_format == "txt":
        return export_txt(segments, enriched=enriched)
    if target_format == "docx":
        return export_docx(segments, enriched=enriched)
    if target_format == "pdf":
        return export_pdf(segments, enriched=enriched)
    raise UnsupportedFormatError(f"Formato de exportación no soportado: {target_format}")


def export_citations_dossier(segments: List[Segment]) -> str:
    """
    Genera un informe textual estructurado (Markdown / texto) con todos los párrafos
    marcados por el usuario para fácil copia y citación académica.
    """
    marked = [s for s in segments if s.is_marked]
    if not marked:
        return "No hay párrafos marcados actualmente."

    lines = [
        "# 📚 FICHA DE CITAS Y PARÁFRASIS SELECCIONADAS",
        f"Total de extractos seleccionados: {len(marked)}\n",
        "---",
    ]

    for i, s in enumerate(marked, start=1):
        lines.append(f"### Cita #{i} — {s.short_provenance}")
        lines.append(f"**📍 Procedencia exacta:** {s.provenance_label}\n")
        lines.append(f"**🌐 Traducción al Español:**\n> \"{s.translated or s.original}\"\n")
        lines.append(f"**📄 Texto Original en Inglés:**\n> \"{s.original}\"\n")
        lines.append("---\n")

    return "\n".join(lines)
