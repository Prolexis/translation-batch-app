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

    # 2. Reparar palabras cortadas por salto de línea con guión tipográfico
    text = re.sub(r'(\b[A-Za-z]+)-\s*\n\s*([a-z]+)', r'\1\2', text)

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

    # Brackets y símbolos matemáticos comunes
    text = text.replace(r'\langle', '<').replace(r'\rangle', '>')
    text = text.replace(r'\cdot', '*').replace(r'\times', 'x')
    text = text.replace(r'\leq', '<=').replace(r'\geq', '>=')
    text = text.replace(r'\neq', '!=').replace(r'\approx', '~')
    text = text.replace(r'\in', ' in ').replace(r'\notin', ' not in ')
    text = text.replace(r'\mathbb{N}', 'N').replace(r'\mathbb{R}', 'R')
    text = text.replace(r'\perp', '[falso]').replace(r'\top', '[verdadero]')
    text = text.replace(r'\triangle', '//')
    text = text.replace(r'\leftarrow', '<-').replace(r'\rightarrow', '->').replace(r'\to', '->')
    text = text.replace(r'\max', 'max').replace(r'\min', 'min').replace(r'\arg', 'arg')
    text = text.replace(r'\sum', 'sum').replace(r'\prod', 'prod')
    text = text.replace(r'\dots', '...').replace(r'\cdots', '...')

    text = re.sub(r'\\text\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\mathrm\{([^{}]+)\}', r'\1', text)
    text = re.sub(r'\\mathbf\{([^{}]+)\}', r'\1', text)

    # Letras griegas LaTeX
    greek_latex = {
        r'\alpha': 'alpha', r'\beta': 'beta', r'\gamma': 'gamma', r'\delta': 'delta',
        r'\epsilon': 'epsilon', r'\varepsilon': 'epsilon', r'\zeta': 'zeta', r'\eta': 'eta',
        r'\theta': 'theta', r'\iota': 'iota', r'\kappa': 'kappa', r'\lambda': 'lambda',
        r'\mu': 'mu', r'\nu': 'nu', r'\xi': 'xi', r'\pi': 'pi', r'\rho': 'rho',
        r'\sigma': 'sigma', r'\tau': 'tau', r'\upsilon': 'upsilon', r'\phi': 'phi',
        r'\chi': 'chi', r'\psi': 'psi', r'\omega': 'omega',
        r'\Gamma': 'Gamma', r'\Delta': 'Delta', r'\Theta': 'Theta', r'\Lambda': 'Lambda',
        r'\Xi': 'Xi', r'\Pi': 'Pi', r'\Sigma': 'Sigma', r'\Phi': 'Phi',
        r'\Psi': 'Psi', r'\Omega': 'Omega'
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
    return False


def _clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = repair_academic_symbols_and_ligatures(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


# Patrones de encabezados y secciones académicas (secciones numéricas limitadas a 1..29 o romanos, excluyendo años como 2035.)
_SECTION_REGEX = re.compile(
    r"^(?:(?:(?:[1-9]|1\d|2\d)(?:\.\d+)*|[IVXLCDM]+)\.?\s+([A-Za-z\u00C0-\u017F\s\-]{2,65})|"
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
    """Resalta visualmente con color las citas in-text [1] y autor-año (Author, Year) adaptables a modo claro y oscuro."""
    if not text:
        return ""
    # Reparar ligaduras y limpiar expresiones matemáticas para lectura fluida
    text = repair_academic_symbols_and_ligatures(text)
    text = clean_math_display(text)

    import html as _html
    safe = _html.escape(text)

    if is_marked:
        # En párrafo marcado (fondo amarillo): texto oscuro de alto contraste
        badge_style = "color: var(--citation-marked-color, #1E1B4B); background: var(--citation-marked-bg, #FDE047); font-weight: 800; padding: 2px 6px; border-radius: 4px; border: 1px solid var(--citation-marked-border, #CA8A04);"
        badge_class = "academic-citation-marked"
    else:
        # En párrafo normal: pastilla con variables CSS adaptables a tema claro y oscuro
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
    if clean_line.lower().startswith("abstract") or clean_line.lower().startswith("resumen"):
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
    Descompone rigurosamente la primera página de un paper científico (IEEE, ACM, ArXiv, etc.)
    en sus componentes estructurales reales:
      (element_type, section_name, text)
      Tipos: metadata, title, authors, affiliations, abstract, keywords, heading, body
    """
    results: List[Tuple[str, Optional[str], str]] = []
    state = "SEARCH_METADATA"
    curr_text: List[str] = []

    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        is_doi_or_rec = bool(re.search(
            r'(?:digital object identifier|doi[\:\s]|10\.\d{4,9}/|received\s+[a-z]+|recibido\s+|accepted\s+|aceptado\s+|date of publication|fecha de publicaci[oó]n|current version)',
            line, re.I
        ))
        is_abs_start = bool(re.match(r'^(?:abstract|resumen)(?:[\:\—\-\.\s]|$)', line, re.I))
        is_key_start = bool(re.match(r'^(?:index terms|keywords|palabras clave|key words|t[eé]rminos de [ií]ndice)(?:[\:\—\-\.\s]|$)', line, re.I))
        is_sec_head = bool(re.match(r'^(?:(?:[IVXLCDM]+|[1-9]\d?)\.|\d+\.\d+)\s+[A-Z\u00C0-\u017F]', line))
        is_author_marker = bool(
            re.search(r'\b(?:dr\.|prof\.|phd|member|fellow|senior member),?\s*(?:ieee)?\b', line, re.I)
            or re.search(r'\b(?:and|y)\s+[A-Z\s\.\-]{3,}\b', line)
            or (line.isupper() and len(line.split()) >= 2 and len(line.split()) <= 12 and not line.endswith('.'))
        )
        is_affil_marker = bool(
            re.search(r'(?:\d+\s*[A-Za-z]*department|\bdepartment|\bdepartamento|\bfaculty|\bfacultad|\bschool|\bescuela|\buniversity|\buniversidad|\binstitute|\binstituto|\blaboratory|\blaboratorio|\bcorresponding author|\bautor de correspondencia|\bthis work was supported|\beste trabajo fue financiado|\bemail[\:\s])', line, re.I)
            or re.match(r'^\d+\s*[A-Z]', line)
        )

        if is_abs_start:
            if curr_text:
                t = state.lower() if state in ["METADATA", "TITLE", "AUTHORS", "AFFILIATIONS"] else "body"
                results.append((t, None, " ".join(curr_text)))
                curr_text = []
            state = "ABSTRACT"
            curr_text.append(line)
            i += 1
            continue

        if is_key_start:
            if curr_text:
                results.append(("abstract" if state == "ABSTRACT" else "body", "Resumen" if state == "ABSTRACT" else None, " ".join(curr_text)))
                curr_text = []
            state = "KEYWORDS"
            curr_text.append(line)
            i += 1
            continue

        if is_sec_head:
            if curr_text:
                results.append((state.lower() if state in ["ABSTRACT", "KEYWORDS"] else "body", None, " ".join(curr_text)))
                curr_text = []
            results.append(("heading", line, line))
            state = "BODY"
            i += 1
            continue

        if state == "SEARCH_METADATA":
            if is_doi_or_rec:
                curr_text.append(line)
                state = "METADATA"
                i += 1
                continue
            else:
                state = "TITLE"
                curr_text.append(line)
                i += 1
                continue

        if state == "METADATA":
            if is_doi_or_rec:
                curr_text.append(line)
                i += 1
                continue
            else:
                results.append(("metadata", "Encabezado", " ".join(curr_text)))
                curr_text = [line]
                state = "TITLE"
                i += 1
                continue

        if state == "TITLE":
            if is_author_marker or is_affil_marker:
                results.append(("title", "Título", " ".join(curr_text)))
                curr_text = [line]
                state = "AUTHORS" if is_author_marker else "AFFILIATIONS"
                i += 1
                continue
            else:
                curr_text.append(line)
                i += 1
                continue

        if state == "AUTHORS":
            if is_affil_marker:
                results.append(("authors", "Autores", " ".join(curr_text)))
                curr_text = [line]
                state = "AFFILIATIONS"
                i += 1
                continue
            else:
                curr_text.append(line)
                i += 1
                continue

        if state == "AFFILIATIONS":
            curr_text.append(line)
            i += 1
            continue

        curr_text.append(line)
        i += 1

    if curr_text:
        t = state.lower() if state in ["METADATA", "TITLE", "AUTHORS", "AFFILIATIONS", "ABSTRACT", "KEYWORDS"] else "body"
        results.append((t, None, " ".join(curr_text)))

    return results


def _parse_body_page_blocks(lines: List[str], current_section: str) -> List[Tuple[str, Optional[str], str]]:
    """Procesa páginas posteriores agrupando párrafos y detectando encabezados y referencias."""
    results: List[Tuple[str, Optional[str], str]] = []
    curr: List[str] = []

    for line in lines:
        is_heading = bool(_SECTION_REGEX.match(line) or _REFERENCES_HEADER_REGEX.match(line))
        is_ref_item = bool(re.match(r"^\[\d+\]", line))
        is_bullet = bool(re.match(r"^[•\-\*]\s+", line))

        if is_heading or is_ref_item or is_bullet:
            if curr:
                clamped = _clamp_paragraph(" ".join(curr))
                for cp in clamped:
                    results.append(("body", current_section, cp))
                curr = []
            if is_heading:
                results.append(("heading", line, line))
            elif is_ref_item:
                results.append(("reference", "Referencias", line))
            continue

        if curr:
            prev_line = curr[-1]
            is_terminator = prev_line.endswith((".", ":", "?", "!"))
            if is_terminator and (line[0].isupper() or line.startswith(("$$\\", "$", "["))):
                clamped = _clamp_paragraph(" ".join(curr))
                for cp in clamped:
                    results.append(("body", current_section, cp))
                curr = [line]
                continue

        curr.append(line)

    if curr:
        clamped = _clamp_paragraph(" ".join(curr))
        for cp in clamped:
            results.append(("body", current_section, cp))

    return results


def _extract_pdf_page_highlights(page) -> List[Tuple[str, str]]:
    """
    Extrae fragmentos de texto anotados con resaltado (/Highlight) o subrayado (/Underline, /Squiggly)
    en una página de PDF. Retorna lista de tuplas (texto_resaltado, color_hex).
    """
    highlights: List[Tuple[str, str]] = []
    annots = getattr(page, "annotations", None)
    if not annots:
        annots = page.get("/Annots")
    if not annots:
        return highlights

    boxes_with_colors = []
    for a in annots:
        try:
            obj = a.get_object() if hasattr(a, "get_object") else a
            if not isinstance(obj, dict):
                continue
            subtype = str(obj.get("/Subtype", ""))
            if subtype not in ("/Highlight", "/Underline", "/Squiggly", "/StrikeOut"):
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

            # Si la anotación incluye texto directo en /Contents o /RC
            contents_text = str(obj.get("/Contents", "") or "").strip()
            if contents_text and len(contents_text) >= 3:
                highlights.append((contents_text, color_hex))

            # Coordenadas: QuadPoints o Rect
            quads = obj.get("/QuadPoints")
            if quads and hasattr(quads, "__iter__"):
                q_coords = [float(x) for x in quads]
                for i in range(0, len(q_coords) - 7, 8):
                    xs = q_coords[i:i+8:2]
                    ys = q_coords[i+1:i+8:2]
                    boxes_with_colors.append((min(xs), min(ys), max(xs), max(ys), color_hex))
            else:
                rect = obj.get("/Rect")
                if rect and hasattr(rect, "__iter__") and len(rect) >= 4:
                    r = [float(x) for x in rect]
                    boxes_with_colors.append((min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3]), color_hex))
        except Exception as e:
            logger.debug("Error procesando anotación de PDF: %s", e)

    if not boxes_with_colors:
        return highlights

    # Extraer texto de la página que cae dentro de las cajas delimitadoras
    box_texts: Dict[int, List[str]] = {i: [] for i in range(len(boxes_with_colors))}

    def visitor(text, cm, tm, font_dict, font_size):
        if not text or not text.strip():
            return
        try:
            x_start = tm[4]
            y = tm[5]
            fsize = font_size if font_size and font_size > 0 else 10.0
            x_end = x_start + len(text) * (fsize * 0.48)
            y_bottom = y - 3
            y_top = y + fsize + 3

            for i, (bx0, by0, bx1, by1, _) in enumerate(boxes_with_colors):
                v_overlap = max(y_bottom, by0) <= min(y_top, by1)
                h_overlap = max(x_start - 4, bx0) <= min(x_end + 4, bx1)
                if v_overlap and h_overlap:
                    box_texts[i].append(text)
                    break
        except Exception:
            pass

    try:
        page.extract_text(visitor_text=visitor)
        for i, (_, _, _, _, col) in enumerate(boxes_with_colors):
            txt = "".join(box_texts[i]).strip()
            if txt and len(txt) >= 2:
                highlights.append((txt, col))
    except Exception as e:
        logger.debug("Error en visitor_text de PDF: %s", e)

    return highlights


def _match_highlight(text_val: str, page_highlights: List[Tuple[str, str]]) -> Tuple[bool, str]:
    """Determina si un texto de párrafo coincide o contiene texto subrayado/resaltado."""
    if not text_val or not page_highlights:
        return False, "#FFD54F"

    clean_val = re.sub(r"\s+", " ", text_val).strip().lower()

    for hl_text, color_hex in page_highlights:
        clean_hl = re.sub(r"\s+", " ", hl_text).strip().lower()
        if len(clean_hl) < 3:
            continue
        if clean_hl in clean_val or clean_val in clean_hl:
            return True, color_hex
        if len(clean_hl) > 20 and clean_hl[:25] in clean_val:
            return True, color_hex
        hl_words = [w for w in clean_hl.split() if len(w) > 4]
        if len(hl_words) >= 3:
            matches = sum(1 for w in hl_words if w in clean_val)
            if matches / len(hl_words) >= 0.65:
                return True, color_hex

    return False, "#FFD54F"


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
        filtered_lines = [
            l for l in lines
            if not _is_running_header_or_footer(l)
            and l not in repeated_headers
            and not re.match(r"^(?:page\s+\d+(?:\s+of\s+\d+)?|\d+|\d+\s*/\s*\d+)$", l, re.I)
        ]

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

            # Comprobar si el texto coincide con alguna anotación de subrayado/resaltado
            is_hl, hl_col = _match_highlight(text_val, page_highlights)

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
            else:
                section_paragraph_counter += 1
                para_num = section_paragraph_counter

            seg = Segment(
                id=seg_id,
                original=text_val,
                section=current_section,
                subsection=current_subsection,
                page=page_idx,
                paragraph_num=para_num,
                element_type=elem_type,
                is_marked=is_hl,
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

        # Detección de texto resaltado o subrayado en runs de DOCX
        is_highlighted_run = False
        run_highlight_color = "#FFD54F"
        if hasattr(p, "runs") and p.runs:
            for r in p.runs:
                if r.font.highlight_color is not None or r.font.underline:
                    is_highlighted_run = True
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
                    break

        seg = Segment(
            id=seg_id,
            original=text,
            section=current_section,
            subsection=current_subsection,
            page=current_page,
            paragraph_num=para_num,
            element_type=elem_type,
            is_marked=is_highlighted_run,
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
        if s.is_marked:
            lines.append(f"★ [SUBRAYADO EN EL ORIGINAL / CITACIÓN]")
            lines.append(text)
            if enriched:
                lines.append(f"   ↳ [Procedencia: {s.provenance_label}]")
        else:
            lines.append(text)
    return "\n\n".join(lines).encode("utf-8")


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
            clean_abs = re.sub(r"^(?:abstract|resumen)\s*[\:\—\-\.]*\s*", "", text, flags=re.I)
            run_text = p.add_run(clean_abs)
            run_text.italic = True
            run_text.font.size = Pt(9.5)
            if ab.is_marked:
                run_text.font.highlight_color = WD_COLOR_INDEX.YELLOW

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
            clean_kw = re.sub(r"^(?:index terms|keywords|palabras clave)\s*[\:\—\-\.]*\s*", "", text, flags=re.I)
            run_kw = p.add_run(clean_kw)
            run_kw.italic = True
            run_kw.font.size = Pt(8.8)

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
        elif elem_type == "reference":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.3)
            p.paragraph_format.first_line_indent = Inches(-0.3)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(text)
            run.font.size = Pt(8)
            if s.is_marked:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15
            run = p.add_run(text)
            run.font.size = Pt(9.5)
            if s.is_marked:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

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

    # 5. Caja de Resumen / Abstract
    if abstract_text:
        y -= 4
        clean_abs = re.sub(r"^(?:abstract|resumen)\s*[\:\—\-\.]*\s*", "", abstract_text, flags=re.I)
        abs_full = "RESUMEN — " + clean_abs
        c.setFont("Helvetica-Oblique", 8.2)
        abs_lines = simpleSplit(abs_full, "Helvetica-Oblique", 8.2, width - 2 * margin - 22)

        kw_lines = []
        if keywords_text:
            clean_kw = re.sub(r"^(?:index terms|keywords|palabras clave)\s*[\:\—\-\.]*\s*", "", keywords_text, flags=re.I)
            kw_full = "PALABRAS CLAVE — " + clean_kw
            kw_lines = simpleSplit(kw_full, "Helvetica-BoldOblique", 7.8, width - 2 * margin - 22)

        box_padding = 8
        abs_box_height = (len(abs_lines) * 11.0) + (len(kw_lines) * 10.2 + 4 if kw_lines else 0) + (2 * box_padding)

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
        for al in abs_lines:
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

        avail_w = col_width - (10 if is_marked else 0)
        lines = simpleSplit(text, font_name, font_size, avail_w)
        if not lines:
            continue

        # 1. Si es encabezado, asegurar que no quede huérfano al final de columna
        if elem_type == "heading":
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

        # 2. Si es párrafo normal o referencia: flujo continuo línea a línea
        curr_y -= space_before
        for line_idx, line in enumerate(lines):
            # Comprobar si cabe la línea en la columna actual
            if curr_y - leading < bottom_margin:
                next_column_or_page()

            col_x = get_col_x(current_col)

            # Fondo y barra ámbar para párrafos marcados/subrayados
            if is_marked:
                c.saveState()
                c.setFillColorRGB(1.0, 0.98, 0.90)  # Amarillo tenue
                c.rect(col_x, curr_y - 2, col_width, leading, fill=1, stroke=0)
                c.setFillColorRGB(0.92, 0.70, 0.12)  # Acento dorado
                c.rect(col_x, curr_y - 2, 3, leading, fill=1, stroke=0)
                c.restoreState()

            c.setFont(font_name, font_size)
            c.setFillColorRGB(*text_color)

            line_x = col_x + (7 if is_marked else 0)
            target_line_w = col_width - (7 if is_marked else 0)
            is_last = (line_idx == len(lines) - 1)

            # Dibujar línea con justificación
            _draw_academic_line(c, line, line_x, curr_y, font_name, font_size, target_line_w, is_last_line=is_last)
            curr_y -= leading

        # Nota de procedencia solo si el párrafo está marcado y es enriquecido
        if show_provenance:
            prov_leading = 9.5
            if curr_y - prov_leading < bottom_margin:
                next_column_or_page()
            col_x = get_col_x(current_col)
            c.setFont("Helvetica-Oblique", 7.2)
            c.setFillColorRGB(0.26, 0.22, 0.75)  # Índigo académico
            c.drawString(col_x + 7, curr_y, f"[📌 Origen: {s.provenance_label}]")
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
