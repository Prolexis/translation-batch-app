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
from dataclasses import dataclass, field
from typing import List, Tuple, Optional

import docx
from docx import Document as DocxDocument
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import Pt, Inches, RGBColor
from pypdf import PdfReader
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit


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
    element_type: str = "body"        # title | abstract | heading | body | caption | reference | formula
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


def _clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
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
    """Resalta visualmente con color las citas in-text [1] y autor-año (Author, Year)."""
    if not text:
        return ""
    import html as _html
    safe = _html.escape(text)

    if is_marked:
        # En párrafo marcado (fondo amarillo): badge azul índigo profundo de alto contraste
        badge_style = "color: #1E1B4B; background: #FDE047; font-weight: 800; padding: 2px 6px; border-radius: 4px; border: 1px solid #CA8A04;"
    else:
        # En párrafo normal: pastilla azul/índigo moderna con borde suave
        badge_style = "color: #818CF8; background: rgba(99, 102, 241, 0.18); font-weight: 700; padding: 2px 6px; border-radius: 4px; border: 1px solid rgba(99, 102, 241, 0.35);"

    def _rep(m):
        return f'<span style="{badge_style}">{m.group(0)}</span>'

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

def _parse_pdf_page_blocks(page_text: str) -> List[str]:
    """Segmenta el texto de una página PDF en párrafos y encabezados reales sin partir oraciones internas."""
    text = page_text.replace("\r\n", "\n").replace("\r", "\n")

    # Si la página ya tiene saltos dobles claros de párrafo, respetarlos
    if "\n\n" in text:
        raw_paras = [p.strip() for p in text.split("\n\n") if p.strip()]
        result = []
        for p in raw_paras:
            lines = [l.strip() for l in p.split("\n") if l.strip()]
            curr = []
            for line in lines:
                is_head = bool(_SECTION_REGEX.match(line) or _REFERENCES_HEADER_REGEX.match(line))
                is_ref = bool(re.match(r"^\[\d+\]", line))
                if is_head or is_ref:
                    if curr:
                        result.append(" ".join(curr))
                        curr = []
                    result.append(line)
                else:
                    curr.append(line)
            if curr:
                result.append(" ".join(curr))
        return result

    # Caso en que el PDF solo tiene saltos simples (\n) por ajuste de línea
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    filtered = []
    for l in lines:
        if re.match(r"^(?:Page\s+\d+\s+of\s+\d+|\d+|Survey\s+on\s+.*)$", l, re.IGNORECASE):
            continue
        filtered.append(l)

    blocks = []
    curr = []
    for line in filtered:
        is_heading = bool(_SECTION_REGEX.match(line) or _REFERENCES_HEADER_REGEX.match(line))
        is_ref_item = bool(re.match(r"^\[\d+\]", line))
        is_bullet = bool(re.match(r"^[•\-\*]\s+", line))

        if is_heading or is_ref_item or is_bullet:
            if curr:
                blocks.append(" ".join(curr))
                curr = []
            blocks.append(line)
            continue

        if curr:
            prev_line = curr[-1]
            # Un fin real de párrafo en PDF termina con puntuación Y la línea se corta antes del margen (longitud < 55)
            is_short_terminator = prev_line.endswith((".", ":", "?", "!")) and len(prev_line) < 55
            if is_short_terminator and (line[0].isupper() or line.startswith(("$$\\", "$", "["))):
                blocks.append(" ".join(curr))
                curr = [line]
                continue

        curr.append(line)

    if curr:
        blocks.append(" ".join(curr))

    return blocks


def extract_academic_pdf(file_bytes: bytes) -> List[Segment]:
    """Extrae párrafos de PDF preservando número de página real y secciones."""
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
    except Exception as exc:
        raise CorruptFileError(f"No se pudo leer el archivo PDF: {exc}") from exc

    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise CorruptFileError("El PDF está protegido con contraseña.") from exc

    segments: List[Segment] = []
    seg_id = 0
    current_section = "Introducción"
    current_subsection = ""
    section_paragraph_counter = 0

    for page_idx, page in enumerate(reader.pages, start=1):
        try:
            raw_text = page.extract_text() or ""
        except Exception:
            raw_text = ""

        raw_text = _clean_text(raw_text)
        if not raw_text:
            continue

        raw_blocks = _parse_pdf_page_blocks(raw_text)

        for b_idx, block in enumerate(raw_blocks):
            norm_block = re.sub(r"[ \t]+", " ", block).strip()
            if not norm_block or len(norm_block) < 3:
                continue

            elem_type, new_sec, new_subsec = _detect_element_type(
                norm_block, current_section, page_idx, is_first=(page_idx == 1 and b_idx == 0)
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
                original=norm_block,
                section=current_section,
                subsection=current_subsection,
                page=page_idx,
                paragraph_num=para_num,
                element_type=elem_type,
            )
            segments.append(seg)
            seg_id += 1

    if not segments:
        raise CorruptFileError("No se pudo extraer texto del PDF (posible PDF escaneado sin capa OCR).")
    return segments


def extract_academic_docx(file_bytes: bytes) -> List[Segment]:
    """Extrae párrafos de DOCX detectando estilos de encabezado y estimando páginas."""
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

        seg = Segment(
            id=seg_id,
            original=text,
            section=current_section,
            subsection=current_subsection,
            page=current_page,
            paragraph_num=para_num,
            element_type=elem_type,
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
        if enriched and s.is_marked:
            lines.append(f"★ [MARCADO / CITACIÓN]")
            lines.append(text)
            lines.append(f"   ↳ [Procedencia: {s.provenance_label}]")
        else:
            lines.append(text)
    return "\n\n".join(lines).encode("utf-8")


def export_docx(segments: List[Segment], enriched: bool = False) -> bytes:
    """
    Exporta a .DOCX con estructura académica fiel al formato del paper:
      - Título centrado y destacado
      - Encabezados con jerarquía (H1, H2)
      - Abstract con sangría y cursiva
      - Referencias bibliográficas formateadas
      - Resaltado y nota de procedencia si enriched=True
    """
    doc = DocxDocument()

    for s in segments:
        text = s.translated or s.original
        elem_type = s.element_type

        if elem_type == "title":
            p = doc.add_heading(text, level=0)
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(14)
        elif elem_type == "heading":
            level = 2 if s.subsection else 1
            p = doc.add_heading(text, level=level)
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(6)
        elif elem_type == "abstract":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.4)
            p.paragraph_format.right_indent = Inches(0.4)
            p.paragraph_format.space_after = Pt(10)
            run = p.add_run(text)
            run.italic = True
            run.font.size = Pt(10)
            if enriched and s.is_marked:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        elif elem_type == "reference":
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.35)
            p.paragraph_format.first_line_indent = Inches(-0.35)
            p.paragraph_format.space_after = Pt(4)
            run = p.add_run(text)
            run.font.size = Pt(8.5)
            if enriched and s.is_marked:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
        else:
            p = doc.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            run = p.add_run(text)
            run.font.size = Pt(10.5)
            if enriched and s.is_marked:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        # Si está marcado y enriquecido, agregar nota de procedencia
        if enriched and s.is_marked and elem_type not in ["title", "heading"]:
            prov_p = doc.add_paragraph()
            prov_p.paragraph_format.left_indent = Inches(0.4)
            prov_p.paragraph_format.space_after = Pt(8)
            prov_run = prov_p.add_run(f"📌 Procedencia: {s.provenance_label}")
            prov_run.italic = True
            prov_run.font.size = Pt(8.5)
            prov_run.font.color.rgb = RGBColor(79, 70, 229)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_pdf(segments: List[Segment], enriched: bool = False) -> bytes:
    """
    Exporta a .PDF con estructura y tipografía académica fiel al paper original:
      - Título en tamaño 15pt bold
      - Encabezados de sección en 12.5pt bold
      - Abstract en 9.5pt cursiva con sangría
      - Cuerpo en 10pt con interlineado óptimo
      - Referencias en 8.5pt
      - Cajas de resaltado y procedencia exacta si enriched=True
    """
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    margin = 2 * cm
    max_width = width - 2 * margin
    y = height - margin

    def check_page_break(needed_space: float):
        nonlocal y
        if y - needed_space < margin:
            c.showPage()
            y = height - margin

    for s in segments:
        text = s.translated or s.original
        elem_type = s.element_type
        is_marked = (enriched and s.is_marked)

        # Configurar estilo según elemento académico
        if elem_type == "title":
            curr_font, curr_size, curr_leading = "Helvetica-Bold", 15, 19
            extra_indent = 0
            spacing_after = 14
        elif elem_type == "heading":
            curr_font, curr_size, curr_leading = "Helvetica-Bold", 12.5, 16
            extra_indent = 0
            spacing_after = 8
        elif elem_type == "abstract":
            curr_font, curr_size, curr_leading = "Helvetica-Oblique", 9.5, 13.5
            extra_indent = 16
            spacing_after = 10
        elif elem_type == "reference":
            curr_font, curr_size, curr_leading = "Helvetica", 8.5, 11
            extra_indent = 8
            spacing_after = 5
        else:
            curr_font, curr_size, curr_leading = "Helvetica", 10, 14
            extra_indent = 0
            spacing_after = 8

        avail_width = max_width - extra_indent - (16 if is_marked else 0)
        lines = simpleSplit(text, curr_font, curr_size, avail_width)
        if not lines:
            continue

        prov_leading = 10
        block_height = len(lines) * curr_leading + (prov_leading + 10 if is_marked else 0) + spacing_after

        check_page_break(block_height)

        # Si está marcado, dibujar recuadro de resaltado
        if is_marked:
            c.saveState()
            c.setFillColorRGB(1.0, 0.96, 0.76)  # Amarillo #FFF3C4
            c.setStrokeColorRGB(0.92, 0.70, 0.10)  # Borde ámbar
            c.roundRect(margin + extra_indent - 4, y - (len(lines) * curr_leading) - 4, avail_width + 8, (len(lines) * curr_leading) + 14, 4, fill=1, stroke=1)
            c.restoreState()

        # Renderizar texto
        c.setFont(curr_font, curr_size)
        c.setFillColorRGB(0.08, 0.08, 0.12)
        for line in lines:
            c.drawString(margin + extra_indent + (6 if is_marked else 0), y, line)
            y -= curr_leading

        # Nota de procedencia si está marcado
        if is_marked:
            y -= 2
            c.setFont("Helvetica-Oblique", 8)
            c.setFillColorRGB(0.26, 0.23, 0.72)
            c.drawString(margin + extra_indent + 8, y, f"[Procedencia: {s.provenance_label}]")
            y -= prov_leading + 4

        y -= spacing_after

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
