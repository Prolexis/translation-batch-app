"""
utils/file_handlers.py
------------------------
Lectura y escritura de archivos en los tres formatos soportados
(.txt, .docx, .pdf), independiente de Streamlit y de los agentes.

Cada extractor devuelve una lista de "segmentos" (párrafos no vacíos, ya
limpios de espacios/rupturas de línea redundantes) que es la unidad atómica
que viajará por el pipeline de agentes.
"""

import io
import re
from dataclasses import dataclass, field
from typing import List

import docx
from pypdf import PdfReader
from docx import Document as DocxDocument
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
    """Unidad mínima de texto (p.ej. un párrafo) que fluye por el pipeline."""
    id: int
    original: str
    translated: str = ""
    status: str = "pendiente"          # pendiente | ok | sospechoso | error
    validation_notes: List[str] = field(default_factory=list)
    retries: int = 0
    color: str = "#FFD54F"


def _clean_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def extract_txt(file_bytes: bytes) -> List[str]:
    try:
        text = file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = file_bytes.decode("latin-1", errors="ignore")
    text = _clean_text(text)
    paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
    if not paragraphs and text:
        paragraphs = [text]
    return paragraphs


def extract_docx(file_bytes: bytes) -> List[str]:
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
    except Exception as exc:
        raise CorruptFileError(f"No se pudo leer el .docx: {exc}") from exc
    paragraphs = []
    for p in doc.paragraphs:
        cleaned = _clean_text(p.text)
        if cleaned:
            paragraphs.append(cleaned)
    if not paragraphs:
        raise CorruptFileError("El documento .docx no contiene texto extraíble.")
    return paragraphs


def extract_pdf(file_bytes: bytes) -> List[str]:
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
    except Exception as exc:
        raise CorruptFileError(f"No se pudo leer el .pdf: {exc}") from exc
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise CorruptFileError("El PDF está protegido/cifrado.") from exc

    paragraphs = []
    for page in reader.pages:
        try:
            page_text = page.extract_text() or ""
        except Exception:
            page_text = ""
        page_text = _clean_text(page_text)
        for block in page_text.split("\n\n"):
            block = block.strip()
            if block:
                paragraphs.append(block)
    if not paragraphs:
        raise CorruptFileError(
            "No se pudo extraer texto del PDF (posible PDF escaneado sin OCR)."
        )
    return paragraphs


def extract_text_segments(filename: str, file_bytes: bytes) -> List[str]:
    """Punto de entrada único usado por el Agente Extractor."""
    ext = "." + filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    if ext == ".txt":
        return extract_txt(file_bytes)
    if ext == ".docx":
        return extract_docx(file_bytes)
    if ext == ".pdf":
        return extract_pdf(file_bytes)
    raise UnsupportedFormatError(f"Formato no soportado: {ext or 'desconocido'}")


# --------------------------------------------------------------------------
# Exportación
# --------------------------------------------------------------------------

def export_txt(segments: List[Segment]) -> bytes:
    text = "\n\n".join(s.translated or s.original for s in segments)
    return text.encode("utf-8")


def export_docx(segments: List[Segment]) -> bytes:
    doc = DocxDocument()
    for s in segments:
        doc.add_paragraph(s.translated or s.original)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_pdf(segments: List[Segment]) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4
    margin = 2 * cm
    max_width = width - 2 * margin
    y = height - margin
    font_name, font_size, leading = "Helvetica", 11, 14

    c.setFont(font_name, font_size)
    for s in segments:
        paragraph = s.translated or s.original
        lines = simpleSplit(paragraph, font_name, font_size, max_width)
        for line in lines:
            if y < margin:
                c.showPage()
                c.setFont(font_name, font_size)
                y = height - margin
            c.drawString(margin, y, line)
            y -= leading
        y -= leading  # espacio entre párrafos
    c.save()
    return buf.getvalue()


def export_segments(segments: List[Segment], target_format: str) -> bytes:
    target_format = target_format.lower().lstrip(".")
    if target_format == "txt":
        return export_txt(segments)
    if target_format == "docx":
        return export_docx(segments)
    if target_format == "pdf":
        return export_pdf(segments)
    raise UnsupportedFormatError(f"Formato de exportación no soportado: {target_format}")
