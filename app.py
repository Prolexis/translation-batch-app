"""
================================================================================
TRADUCTOR DE PAPERS ACADÉMICOS CON MARCADO Y TRAZABILIDAD DE ORIGEN
================================================================================
Aplicación web en Streamlit con agentes LangChain y Google Gemini (Flash-Lite)
para la traducción de artículos científicos y papers en inglés (.pdf, .docx, .txt)
al español, con trazabilidad exacta de origen:
  - Detección y preservación de sección, subsección, página y número de párrafo
  - Previsualización inmediata de la traducción y reportes ejecutivos
  - Selección interactiva de párrafos con notificación visual instantánea
  - Marcado robusto y persistente para citación y canasta de extractos
  - Exportación inmediata y enriquecida con párrafos resaltados y notas al pie
================================================================================
"""

import io
import os
import re
import html
import time
import base64
import zipfile
import logging
from typing import List, Set, Optional, Dict, Any, Tuple

import streamlit as st

from config import settings
from utils.file_handlers import (
    Segment,
    extract_academic_segments,
    export_segments,
    export_citations_dossier,
    highlight_citations_html,
    UnsupportedFormatError,
    CorruptFileError,
)
from utils.colors import color_for_index, render_highlighted_html
from agents.orchestrator import TranslationOrchestrator

# Configuración de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("translation_app.ui")

# Configuración de página en Streamlit
st.set_page_config(
    page_title="Traductor de Papers con Trazabilidad de Origen",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inyección de estilos CSS avanzados
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;500;600;700;800&family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Outfit', sans-serif !important;
        letter-spacing: -0.02em;
    }

    /* Hero Glassmorphic Card */
    .hero-banner {
        background: linear-gradient(135deg, rgba(15, 23, 42, 0.88) 0%, rgba(30, 41, 59, 0.78) 100%);
        border: 1px solid rgba(99, 102, 241, 0.35);
        border-radius: 18px;
        padding: 26px 32px;
        margin-bottom: 24px;
        box-shadow: 0 12px 36px 0 rgba(0, 0, 0, 0.38);
        backdrop-filter: blur(14px);
    }

    .hero-title {
        font-size: 2.25rem;
        font-weight: 800;
        background: linear-gradient(135deg, #818CF8 0%, #C084FC 45%, #38BDF8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
    }

    .agent-pipeline {
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
        margin-top: 14px;
    }

    .agent-pill {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        padding: 6px 14px;
        background: rgba(255, 255, 255, 0.06);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 30px;
        font-size: 0.82rem;
        font-weight: 600;
        color: #E2E8F0;
        transition: all 0.2s ease;
    }

    /* Provenance Box Standard */
    .provenance-card {
        background: linear-gradient(135deg, #1E1B4B 0%, #312E81 100%);
        border: 1px solid #6366F1;
        border-left: 6px solid #818CF8;
        border-radius: 12px;
        padding: 16px 20px;
        margin: 12px 0;
        box-shadow: 0 6px 18px rgba(0, 0, 0, 0.25);
    }

    /* Provenance Box Marked (Glowing Amber) */
    .provenance-card-marked {
        background: linear-gradient(135deg, rgba(120, 53, 15, 0.45) 0%, rgba(69, 26, 3, 0.75) 100%);
        border: 2px solid #F59E0B;
        border-left: 8px solid #FACC15;
        border-radius: 12px;
        padding: 18px 22px;
        margin: 12px 0;
        box-shadow: 0 8px 24px rgba(245, 158, 11, 0.28);
    }

    .provenance-title {
        font-size: 0.82rem;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        color: #A5B4FC;
        font-weight: 700;
        margin-bottom: 4px;
    }

    .provenance-body {
        font-size: 1.15rem;
        font-weight: 700;
        color: #FFFFFF;
    }

    /* Marked paragraph card */
    .paragraph-marked {
        background: rgba(254, 240, 138, 0.16) !important;
        border: 2px solid #FACC15 !important;
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 12px;
        box-shadow: 0 4px 14px rgba(250, 204, 21, 0.18);
        transition: all 0.2s ease;
    }

    .paragraph-normal {
        background: rgba(30, 41, 59, 0.4);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 10px;
        padding: 14px 18px;
        margin-bottom: 12px;
        transition: all 0.2s ease;
    }

    .citation-badge {
        display: inline-block;
        padding: 3px 8px;
        border-radius: 6px;
        font-size: 0.75rem;
        font-weight: 600;
        background: #4F46E5;
        color: #FFFFFF;
        margin-bottom: 6px;
    }

    .citation-badge-marked {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 0.78rem;
        font-weight: 800;
        background: #F59E0B;
        color: #000000;
        margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# Estado de la Sesión
# ------------------------------------------------------------------------------
def init_session_state():
    st.session_state.setdefault("files_status", {})
    st.session_state.setdefault("results", {})
    st.session_state.setdefault("uploaded_map", {})
    st.session_state.setdefault("processing", False)
    st.session_state.setdefault("selected_paragraph_idx", 0)
    st.session_state.setdefault("active_file", "")
    st.session_state.setdefault("marked_ids", set())  # Set de enteros con IDs marcados


init_session_state()

# ------------------------------------------------------------------------------
# Barra Lateral
# ------------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### ⚙️ Configuración del Traductor")

    env_api_key = settings.GEMINI_API_KEY.strip()
    if env_api_key:
        st.success("🔒 GEMINI_API_KEY activa (.env)")
        with st.expander("🔑 Cambiar API Key"):
            override_key = st.text_input(
                "Nueva Clave API (opcional)",
                value="",
                type="password",
                help="Deja en blanco para usar la clave del .env.",
            )
        api_key_input = override_key.strip() if override_key.strip() else env_api_key
    else:
        api_key_input = st.text_input(
            "GEMINI_API_KEY (Obligatoria)",
            value="",
            type="password",
            help="Ingresa tu clave de Gemini.",
        )
        if not api_key_input:
            st.warning("⚠️ Debes proporcionar tu GEMINI_API_KEY.")

    st.markdown("---")
    st.markdown("#### 🌐 Parámetros de Traducción")

    col_s, col_t = st.columns(2)
    with col_s:
        source_lang = st.selectbox(
            "Origen",
            options=["en", "auto", "fr", "de"],
            index=0,
            format_func=lambda x: {"en": "🇬🇧 Inglés", "auto": "🌐 Auto", "fr": "🇫🇷 Francés", "de": "🇩🇪 Alemán"}.get(x, x),
        )
    with col_t:
        target_lang = st.selectbox(
            "Destino",
            options=["es", "en", "fr", "pt"],
            index=0,
            format_func=lambda x: {"es": "🇪🇸 Español", "en": "🇬🇧 Inglés", "fr": "🇫🇷 Francés", "pt": "🇵🇹 Portugués"}.get(x, x),
        )

    alignment_mode = st.selectbox(
        "Alineación",
        options=["position", "semantic"],
        index=0,
        format_func=lambda x: "Rápida por posición" if x == "position" else "Semántica (Embeddings)",
    )

    st.markdown("---")
    st.markdown("#### 🧪 Paper de Prueba (5+ Páginas)")
    st.caption("Carga con 1 clic el paper científico de muestra con 5 páginas, citas [1] y fórmulas:")

    sample_col1, sample_col2 = st.columns(2)
    sample_dir = os.path.join(os.path.dirname(__file__), "sample_files")

    with sample_col1:
        if st.button("📄 Probar PDF (5 págs)", use_container_width=True):
            pdf_path = os.path.join(sample_dir, "academic_paper_sample.pdf")
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    st.session_state.uploaded_map["academic_paper_sample.pdf"] = f.read()
                    st.session_state.files_status["academic_paper_sample.pdf"] = "pendiente"
                    st.session_state.active_file = "academic_paper_sample.pdf"
                st.session_state.marked_ids.clear()
                st.success("Paper PDF cargado.")
                st.rerun()
            else:
                st.error("No se encontró el archivo de muestra PDF.")

    with sample_col2:
        if st.button("📝 Probar DOCX (5 págs)", use_container_width=True):
            docx_path = os.path.join(sample_dir, "academic_paper_sample.docx")
            if os.path.exists(docx_path):
                with open(docx_path, "rb") as f:
                    st.session_state.uploaded_map["academic_paper_sample.docx"] = f.read()
                    st.session_state.files_status["academic_paper_sample.docx"] = "pendiente"
                    st.session_state.active_file = "academic_paper_sample.docx"
                st.session_state.marked_ids.clear()
                st.success("Paper DOCX cargado.")
                st.rerun()
            else:
                st.error("No se encontró el archivo de muestra DOCX.")

    st.markdown("---")
    st.markdown("#### 📖 Documentación de Arquitectura")
    with st.expander("ℹ️ Ver Ciclo de Vida de Trazabilidad"):
        doc_path = os.path.join(os.path.dirname(__file__), "docs", "trazabilidad_explicacion.md")
        if os.path.exists(doc_path):
            with open(doc_path, "r", encoding="utf-8") as f:
                st.markdown(f.read())

# ------------------------------------------------------------------------------
# Encabezado Principal (Hero)
# ------------------------------------------------------------------------------
st.markdown("""
<div class="hero-banner">
    <div class="hero-title">🎓 Traductor de Papers Académicos con Trazabilidad y Marcado de Origen</div>
    <div style="color: #CBD5E1; font-size: 1.02rem; line-height: 1.55;">
        Sube un artículo o paper científico en inglés (<strong>PDF, DOCX o TXT</strong>) y el sistema lo traducirá 
        completamente al español mediante agentes coordinados con la API de <strong>Gemini 3.1 Flash-Lite</strong>.
        Descarga de inmediato la traducción completa o selecciona cualquier párrafo para auditar su 
        <strong>sección, página y número de párrafo exacto</strong> en el documento original.
    </div>
    <div class="agent-pipeline">
        <span class="agent-pill">📄 1. Extractor Estructurado (Páginas y Secciones)</span>
        <span class="agent-pill">🌐 2. Traductor Académico en Paralelo (Citas [1] y Fórmulas)</span>
        <span class="agent-pill">🛡️ 3. Validador de Calidad</span>
        <span class="agent-pill">🔗 4. Trazabilidad de Origen y Previsualización</span>
    </div>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# 1. Carga de Archivos
# ------------------------------------------------------------------------------
uploaded_files = st.file_uploader(
    "📁 Sube aquí tu paper o artículo académico en inglés:",
    type=["pdf", "docx", "txt"],
    accept_multiple_files=True,
    help="Soporta papers de cualquier longitud con secciones numeradas, citas y tablas.",
)

if uploaded_files:
    for f in uploaded_files:
        if f.name not in st.session_state.uploaded_map:
            st.session_state.uploaded_map[f.name] = f.getvalue()
            st.session_state.files_status.setdefault(f.name, "pendiente")
            if not st.session_state.active_file:
                st.session_state.active_file = f.name

if st.session_state.uploaded_map:
    file_list = list(st.session_state.uploaded_map.keys())
    if not st.session_state.active_file or st.session_state.active_file not in file_list:
        st.session_state.active_file = file_list[0]

    btn_col1, btn_col2, btn_col3 = st.columns([3, 2, 2])
    with btn_col1:
        start_button = st.button(
            "🚀 Traducir Paper Académico Completo",
            type="primary",
            use_container_width=True,
            disabled=st.session_state.processing,
        )
    with btn_col2:
        if len(file_list) > 1:
            st.session_state.active_file = st.selectbox(
                "Documento activo para revisar:",
                options=file_list,
                index=file_list.index(st.session_state.active_file),
            )
        else:
            st.caption(f"📄 **Archivo cargado:** `{file_list[0]}`")
    with btn_col3:
        if st.button("🗑️ Limpiar / Nuevo Documento", use_container_width=True, disabled=st.session_state.processing):
            st.session_state.files_status.clear()
            st.session_state.results.clear()
            st.session_state.uploaded_map.clear()
            st.session_state.marked_ids.clear()
            st.session_state.active_file = ""
            st.rerun()

    # Procesamiento Automático
    if start_button:
        if not api_key_input:
            st.error("⚠️ Debes proporcionar la GEMINI_API_KEY antes de procesar.")
        else:
            st.session_state.processing = True
            st.session_state.marked_ids.clear()
            orchestrator = TranslationOrchestrator(
                source_lang=source_lang,
                target_lang=target_lang,
                alignment_mode=alignment_mode,
                api_key=api_key_input,
            )

            progress_bar = st.progress(0.0, text="Iniciando traducción académica...")
            files = list(st.session_state.uploaded_map.keys())
            total_files = len(files)

            for idx, fname in enumerate(files):
                st.session_state.files_status[fname] = "procesando"

                def _progress_cb(stage_msg: str, frac: float, _f=fname, _i=idx):
                    overall = (_i + frac) / total_files
                    progress_bar.progress(min(overall, 1.0), text=f"[{_f}] {stage_msg}")

                try:
                    f_bytes = st.session_state.uploaded_map[fname]
                    ctx = orchestrator.process_file(fname, f_bytes, on_progress=_progress_cb)
                except Exception as exc:  # noqa: BLE001
                    logger.exception("Error procesando paper %s", fname)
                    ctx = {"error": str(exc), "file_status": "error", "segments": []}

                st.session_state.results[fname] = ctx
                st.session_state.files_status[fname] = ctx.get("file_status", "error")

            progress_bar.progress(1.0, text="Traducción completada con éxito.")
            st.session_state.processing = False
            st.rerun()

# ------------------------------------------------------------------------------
# 2. Resultados, Previsualización y Reportes
# ------------------------------------------------------------------------------
active_fname = st.session_state.active_file
active_context = st.session_state.results.get(active_fname)

if active_context and active_context.get("segments"):
    segments: List[Segment] = active_context["segments"]
    total_segs = len(segments)

    # Sincronizar estado de marcado con el session_state persistente
    for s in segments:
        s.is_marked = (s.id in st.session_state.marked_ids)

    marked_segs = [s for s in segments if s.is_marked]

    st.markdown("---")
    st.markdown("### 📄 1. Previsualización del Paper Traducido y Descarga Inmediata")
    st.caption("Examina el documento maquetado tal cual el paper original antes de descargarlo en tu formato preferido:")

    # Generar binarios de exportación
    pdf_bytes = export_segments(segments, "pdf", enriched=False)
    docx_bytes = export_segments(segments, "docx", enriched=False)
    txt_bytes = export_segments(segments, "txt", enriched=False)
    b64_pdf = base64.b64encode(pdf_bytes).decode("utf-8")

    # Botonera de Descarga Destacada
    dl_col1, dl_col2, dl_col3, dl_col4 = st.columns([3, 3, 2, 2])
    with dl_col1:
        st.download_button(
            label="⬇️ Descargar .PDF (2 Col. IEEE)",
            data=pdf_bytes,
            file_name=f"{active_fname.rsplit('.', 1)[0]}_traducido.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary",
        )
    with dl_col2:
        st.download_button(
            label="⬇️ Descargar .DOCX (2 Columnas)",
            data=docx_bytes,
            file_name=f"{active_fname.rsplit('.', 1)[0]}_traducido.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary",
        )
    with dl_col3:
        st.download_button(
            label="⬇️ Descargar .TXT",
            data=txt_bytes,
            file_name=f"{active_fname.rsplit('.', 1)[0]}_traducido.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with dl_col4:
        st.metric("Párrafos", total_segs, help="Total de párrafos académicos estructurados")

    # Contenedor de Previsualización Interactiva antes de descargar
    with st.container(border=True):
        p_mode = st.radio(
            "Modo de previsualización antes de descargar:",
            options=["📑 Visor PDF Interactivo (Maquetación Real en 2 Columnas)", "📜 Hoja de Paper Científico (Lectura Rápida con Citas Coloreadas)"],
            horizontal=True,
            key="pre_download_preview_selector",
        )

        if p_mode.startswith("📑 Visor PDF"):
            st.caption("Visualiza el PDF generado directamente en tu navegador con maquetación de dos columnas, encabezados y abstract:")
            pdf_embed_html = (
                f'<iframe src="data:application/pdf;base64,{b64_pdf}#toolbar=1&navpanes=0" '
                f'width="100%" height="680px" '
                f'style="border-radius:10px; border:1px solid rgba(255,255,255,0.15); box-shadow: 0 10px 30px rgba(0,0,0,0.35); background:#1E293B;">'
                f'</iframe>'
            )
            st.markdown(pdf_embed_html, unsafe_allow_html=True)
        else:
            # Hoja de Paper Científico (Simulador de Paper Real en HTML/CSS)
            st.caption("Previsualización estructurada con título, abstract enmarcado y citas resaltadas con color:")

            title_txt = next((s.translated or s.original for s in segments if s.element_type == "title"), segments[0].translated or segments[0].original)
            abstract_txt = " ".join(s.translated or s.original for s in segments if s.element_type == "abstract" or s.section.lower() in ["abstract", "resumen"])

            paper_sheet_html = [
                '<div style="background: #0B0F19; border: 1px solid #1E293B; border-radius: 12px; padding: 28px 36px; box-shadow: 0 15px 35px rgba(0,0,0,0.4); max-height: 600px; overflow-y: auto;">',
                '<div style="text-align: center; font-size: 0.75rem; letter-spacing: 1.5px; color: #64748B; font-weight: 700; text-transform: uppercase; margin-bottom: 12px; border-bottom: 1px solid #1E293B; padding-bottom: 8px;">REVISTA CIENTÍFICA · TRADUCCIÓN ACADÉMICA IEEE</div>',
                f'<h2 style="text-align: center; color: #F8FAFC; font-size: 1.45rem; font-weight: 800; line-height: 1.35; margin: 16px 0 8px 0;">{html.escape(title_txt)}</h2>',
                '<div style="text-align: center; color: #94A3B8; font-size: 0.85rem; margin-bottom: 20px;">Artículo Académico Traducido con Preservación de Citas Bibliográficas y Trazabilidad</div>',
            ]

            if abstract_txt:
                clean_abs = re.sub(r"^(?:abstract|resumen)\s*[\:\—\-\.]*\s*", "", abstract_txt, flags=re.I)
                abs_colored = highlight_citations_html(clean_abs, is_marked=False)
                paper_sheet_html.append(
                    f'<div style="background: rgba(30, 41, 59, 0.7); border: 1px solid #334155; border-radius: 8px; padding: 14px 18px; margin-bottom: 24px;">'
                    f'<strong style="color: #60A5FA; font-size: 0.88rem; letter-spacing: 0.5px;">RESUMEN — </strong>'
                    f'<span style="font-style: italic; color: #CBD5E1; font-size: 0.90rem; line-height: 1.6;">{abs_colored}</span>'
                    f'</div>'
                )

            paper_sheet_html.append('<div style="border-top: 1px solid #1E293B; padding-top: 16px;">')
            for s in segments[:18]:
                if s.element_type in ["title", "abstract"] or s.section.lower() in ["abstract", "resumen"]:
                    continue
                if s.element_type == "heading":
                    paper_sheet_html.append(f'<h4 style="color: #93C5FD; margin-top: 16px; margin-bottom: 6px; font-size: 1.05rem; text-transform: uppercase;">{html.escape(s.translated or s.original)}</h4>')
                elif s.element_type == "reference":
                    paper_sheet_html.append(f'<div style="font-size: 0.78rem; color: #64748B; margin-bottom: 4px; padding-left: 14px; text-indent: -14px;">{html.escape(s.translated or s.original)}</div>')
                else:
                    c_txt = highlight_citations_html(s.translated or s.original, is_marked=s.is_marked)
                    paper_sheet_html.append(f'<p style="color: #E2E8F0; font-size: 0.90rem; line-height: 1.65; margin-bottom: 12px;">{c_txt}</p>')

            if len(segments) > 18:
                paper_sheet_html.append(f'<div style="text-align: center; color: #64748B; font-size: 0.82rem; margin-top: 16px;">... y {len(segments) - 18} párrafos más. Explóralos todos o marca citas en la sección 2 a continuación.</div>')

            paper_sheet_html.append('</div></div>')
            st.markdown("".join(paper_sheet_html), unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # 2. Trazabilidad de Origen, Marcado y Reportes
    # --------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 🔬 2. Trazabilidad de Origen, Marcado para Citas y Auditoría")
    st.caption(
        "Revisa la correspondencia párrafo a párrafo, audita su procedencia exacta (sección, página y número) "
        "y selecciona extractos para generar tu ficha de citación y versiones enriquecidas:"
    )

    # Callbacks de marcado nativos de Streamlit (actualizan estado antes del ciclo de renderizado)
    def _toggle_mark_seg(seg_id: int, p_idx: Optional[int] = None):
        if seg_id in st.session_state.marked_ids:
            st.session_state.marked_ids.discard(seg_id)
        else:
            st.session_state.marked_ids.add(seg_id)
        if p_idx is not None:
            st.session_state.selected_paragraph_idx = p_idx

    # Pestañas con títulos estáticos para evitar que React desmonte el árbol al marcar
    tab_preview, tab_inspector, tab_reports, tab_citations = st.tabs([
        "📖 Previsualización del Paper Traducido",
        "🔬 Inspector de Procedencia y Edición",
        "📊 Reporte Ejecutivo de Traducción",
        "📚 Canasta de Citas Seleccionadas",
    ])

    # --------------------------------------------------------------------------
    # Pestaña 1: Previsualización de la Traducción
    # --------------------------------------------------------------------------
    with tab_preview:
        view_mode = st.radio(
            "Selecciona el modo de previsualización:",
            options=["Lienzo de Lectura Académica", "Vista Bilingüe Sincronizada (Lado a Lado)"],
            horizontal=True,
        )

        if view_mode == "Lienzo de Lectura Académica":
            st.caption("Lectura estructurada del paper traducido. Las citas [1] y (Autor, Año) se resaltan con color, y los párrafos marcados se destacan en amarillo con su nota de origen.")

            # Paginación fluida para documentos extensos (evita sobrecarga del DOM y permite navegación instantánea)
            PAGE_SIZE = 25
            total_pages = max(1, (total_segs + PAGE_SIZE - 1) // PAGE_SIZE)
            if total_pages > 1:
                pag_c1, pag_c2, pag_c3 = st.columns([2, 5, 2])
                with pag_c2:
                    current_preview_page = st.selectbox(
                        "Bloque de lectura del paper:",
                        options=list(range(1, total_pages + 1)),
                        format_func=lambda p: f"Página {p} de {total_pages} (Párrafos {(p-1)*PAGE_SIZE + 1} - {min(p*PAGE_SIZE, total_segs)} de {total_segs})",
                        key="preview_page_selector",
                    )
                start_p_idx = (current_preview_page - 1) * PAGE_SIZE
                end_p_idx = min(start_p_idx + PAGE_SIZE, total_segs)
                visible_segments = list(enumerate(segments[start_p_idx:end_p_idx], start=start_p_idx))
            else:
                visible_segments = list(enumerate(segments))

            curr_sec = ""
            for i, s in visible_segments:
                if s.section != curr_sec and s.section:
                    curr_sec = s.section
                    st.markdown(f"### {curr_sec}")

                is_m = (s.id in st.session_state.marked_ids)
                text_to_show = s.translated or s.original
                colored_html = highlight_citations_html(text_to_show, is_marked=is_m)

                # Tarjeta de párrafo con estructura HTML uniforme y estable (previene errores de reconciliación React)
                with st.container(border=True):
                    p_col1, p_col2 = st.columns([10, 2])
                    with p_col1:
                        if is_m:
                            header_badge = f'<div style="background:#FDE047; color:#1E1B4B; font-weight:800; font-size:0.75rem; padding:3px 8px; border-radius:4px; display:inline-block; margin-bottom:8px;">⭐ MARCADO PARA CITACIÓN · {html.escape(s.short_provenance)}</div>'
                            prov_foot = f'<div style="margin-top:8px; font-size:0.80rem; color:#FDE047; font-style:italic;">📌 {html.escape(s.provenance_label)}</div>'
                            content_style = "font-size: 1.02rem; line-height: 1.65; color: #FEF08A;"
                        else:
                            header_badge = f'<div style="color:#94A3B8; font-size:0.75rem; margin-bottom:6px;">🏷️ {html.escape(s.short_provenance)}</div>'
                            prov_foot = ""
                            content_style = "font-size: 0.98rem; line-height: 1.6; color: #F1F5F9;"

                        st.markdown(
                            f'{header_badge}'
                            f'<div style="{content_style}">{colored_html}</div>'
                            f'{prov_foot}',
                            unsafe_allow_html=True,
                        )

                    with p_col2:
                        btn_text = "❌ Quitar" if is_m else "⭐ Marcar"
                        btn_type = "secondary" if is_m else "primary"
                        st.button(
                            btn_text,
                            key=f"btn_prev_mark_{s.id}",
                            on_click=_toggle_mark_seg,
                            args=(s.id, i),
                            use_container_width=True,
                            type=btn_type,
                        )

        else:
            # Vista Bilingüe Sincronizada (Lado a Lado)
            st.caption("Verificación cruzada lado a lado: correspondencia exacta por párrafos con citas coloreadas.")
            col_left, col_right = st.columns(2)
            with col_left:
                st.markdown(f"#### 📄 Documento Original ({source_lang.upper()})")
                with st.container(height=520):
                    left_html_blocks = []
                    for s in segments:
                        is_m = (s.id in st.session_state.marked_ids)
                        bg = "rgba(254, 240, 138, 0.18)" if is_m else "rgba(255, 255, 255, 0.04)"
                        fg = "#FEF08A" if is_m else "#E2E8F0"
                        border_color = "#F59E0B" if is_m else "#6366F1"
                        star = "⭐ " if is_m else ""
                        colored_orig = highlight_citations_html(s.original, is_marked=is_m)
                        left_html_blocks.append(
                            f'<div style="background:{bg}; color:{fg}; padding:10px 14px; border-radius:8px; margin-bottom:10px; font-size:0.92rem; border-left: 4px solid {border_color};">'
                            f'<strong style="font-size:0.75rem;">{star}[{html.escape(s.short_provenance)}]</strong><br>'
                            f'{colored_orig}</div>'
                        )
                    st.markdown("\n".join(left_html_blocks), unsafe_allow_html=True)

            with col_right:
                st.markdown(f"#### 🌐 Traducción Generada ({target_lang.upper()})")
                with st.container(height=520):
                    right_html_blocks = []
                    for s in segments:
                        is_m = (s.id in st.session_state.marked_ids)
                        bg = "rgba(254, 240, 138, 0.18)" if is_m else "rgba(255, 255, 255, 0.04)"
                        fg = "#FEF08A" if is_m else "#E2E8F0"
                        border_color = "#F59E0B" if is_m else "#10B981"
                        star = "⭐ " if is_m else ""
                        colored_trans = highlight_citations_html(s.translated or s.original, is_marked=is_m)
                        right_html_blocks.append(
                            f'<div style="background:{bg}; color:{fg}; padding:10px 14px; border-radius:8px; margin-bottom:10px; font-size:0.92rem; border-left: 4px solid {border_color};">'
                            f'<strong style="font-size:0.75rem;">{star}[{html.escape(s.short_provenance)}]</strong><br>'
                            f'{colored_trans}</div>'
                        )
                    st.markdown("\n".join(right_html_blocks), unsafe_allow_html=True)

    # --------------------------------------------------------------------------
    # Pestaña 2: Inspector de Procedencia y Edición Párrafo a Párrafo
    # --------------------------------------------------------------------------
    with tab_inspector:
        def _format_option(i: int) -> str:
            s = segments[i]
            prefix = "⭐ [MARCADO] " if (s.id in st.session_state.marked_ids) else ""
            snip = s.translated[:65] if s.translated else s.original[:65]
            return f"{prefix}P{i+1}: Pág. {s.page} | {s.section[:22]} ➔ {snip}..."

        current_idx = st.selectbox(
            "Selecciona un párrafo para auditar su origen y editarlo:",
            options=list(range(total_segs)),
            index=min(st.session_state.selected_paragraph_idx, total_segs - 1),
            format_func=_format_option,
        )
        st.session_state.selected_paragraph_idx = current_idx
        sel_seg = segments[current_idx]
        is_sel_marked = (sel_seg.id in st.session_state.marked_ids)

        # Ficha Destacada de Procedencia con estructura estable
        with st.container(border=True):
            if is_sel_marked:
                st.markdown('<div style="background:#FDE047; color:#1E1B4B; font-weight:800; font-size:0.80rem; padding:4px 10px; border-radius:4px; display:inline-block; margin-bottom:8px;">⭐ PÁRRAFO MARCADO Y REGISTRADO PARA CITACIÓN / AUDITORÍA</div>', unsafe_allow_html=True)
            st.markdown(f"#### 📍 {sel_seg.provenance_label}")
            st.caption(
                f"**Identificador:** `seg_id #{sel_seg.id}` &nbsp;|&nbsp; "
                f"**Sección:** `{sel_seg.section}` &nbsp;|&nbsp; "
                f"**Página:** `{sel_seg.page}` &nbsp;|&nbsp; "
                f"**Párrafo:** `{sel_seg.paragraph_num}` &nbsp;|&nbsp; "
                f"**Tipo:** `{sel_seg.element_type}`"
            )

        # Botonera de Marcado y Navegación
        btn_mark_col, btn_nav_col = st.columns([3, 3])
        with btn_mark_col:
            mark_btn_label = "❌ Quitar Selección (Desmarcar)" if is_sel_marked else "⭐ Marcar este Párrafo para Cita / Tesis"
            st.button(
                mark_btn_label,
                key=f"btn_inspector_mark_{sel_seg.id}",
                on_click=_toggle_mark_seg,
                args=(sel_seg.id,),
                use_container_width=True,
                type="secondary" if is_sel_marked else "primary",
            )

        with btn_nav_col:
            nav1, nav2 = st.columns(2)
            with nav1:
                if st.button("⬅️ Párrafo Anterior", disabled=(current_idx == 0), use_container_width=True):
                    st.session_state.selected_paragraph_idx = max(0, current_idx - 1)
                    st.rerun()
            with nav2:
                if st.button("Párrafo Siguiente ➡️", disabled=(current_idx >= total_segs - 1), use_container_width=True):
                    st.session_state.selected_paragraph_idx = min(total_segs - 1, current_idx + 1)
                    st.rerun()


        # Vista Paralela Bilingüe
        col_orig, col_trans = st.columns(2)
        with col_orig:
            st.markdown("#### 📄 Texto Original en Inglés (Fuente)")
            st.text_area(
                "Original",
                value=sel_seg.original,
                height=180,
                disabled=True,
                key=f"orig_area_{sel_seg.id}",
            )
            st.caption(f"Longitud: {len(sel_seg.original.split())} palabras · Página {sel_seg.page}")

        with col_trans:
            st.markdown("#### 🌐 Traducción al Español (Gemini)")
            edited_val = st.text_area(
                "Traducción",
                value=sel_seg.translated or sel_seg.original,
                height=180,
                key=f"trans_area_{sel_seg.id}",
                help="Puedes ajustar la redacción si lo deseas antes de descargar.",
            )
            sel_seg.translated = edited_val

            if sel_seg.validation_notes:
                st.caption(f"🛡️ Notas de validación: {'; '.join(sel_seg.validation_notes)}")
            else:
                st.caption("🛡️ Integridad de citas y formato: OK")

    # --------------------------------------------------------------------------
    # Pestaña 3: Reporte Ejecutivo de Traducción (Resumido y Claro)
    # --------------------------------------------------------------------------
    with tab_reports:
        st.markdown("#### 📊 Reporte Ejecutivo de la Traducción")
        st.caption("Resumen consolidado de métricas clave y validación académica.")

        rep_m1, rep_m2, rep_m3, rep_m4 = st.columns(4)
        refs_count = sum(1 for s in segments if s.element_type == "reference")
        forms_count = sum(1 for s in segments if s.element_type == "formula")
        total_words = sum(len((s.translated or s.original).split()) for s in segments)
        duration = active_context.get("duration_seconds", 0.0)

        with rep_m1:
            st.metric("Total Párrafos", total_segs)
        with rep_m2:
            st.metric("Palabras Traducidas", f"{total_words:,}")
        with rep_m3:
            st.metric("Referencias Protegidas", refs_count, help="Sección bibliográfica preservada intacta")
        with rep_m4:
            st.metric("Tiempo Total", f"{duration:.1f}s" if duration else "Rápido")

        st.markdown("---")
        # Tarjeta de Estado Resumida
        st.success(
            f"✅ **Traducción Finalizada con Éxito** — "
            f"`{active_fname}` ({source_lang.upper()} ➔ {target_lang.upper()}) · "
            f"Citas bibliográficas in-text preservadas · "
            f"{len(st.session_state.marked_ids)} extractos seleccionados para citación."
        )

        with st.expander("🔍 Ver Registro Técnico de Agentes (Opcional)"):
            st.caption("Eventos registrados por los agentes durante el pipeline:")
            for log_entry in active_context.get("memory_log", []):
                st.text(f"• {log_entry}")

    # --------------------------------------------------------------------------
    # Pestaña 4: Canasta de Citas Seleccionadas
    # --------------------------------------------------------------------------
    with tab_citations:
        if not marked_segs:
            st.info("Aún no has marcado ningún párrafo. Puedes marcar párrafos en la pestaña de **Previsualización** o en el **Inspector** para generar tu ficha de citación.")
        else:
            st.success(f"Has seleccionado **{len(marked_segs)} extractos** con trazabilidad de origen.")

            dossier_text = export_citations_dossier(segments)

            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.markdown("#### 📋 Ficha de Citas Generada")
            with d_col2:
                st.download_button(
                    label="⬇️ Descargar Ficha (.TXT)",
                    data=dossier_text.encode("utf-8"),
                    file_name=f"{active_fname.rsplit('.', 1)[0]}_citas_seleccionadas.txt",
                    mime="text/plain",
                    use_container_width=True,
                )

            st.code(dossier_text, language="markdown")

    # --------------------------------------------------------------------------
    # 4. Descargas Enriquecidas con Trazabilidad (Resumido y directo)
    # --------------------------------------------------------------------------
    st.markdown("---")
    st.markdown("### 🌟 3. Descarga Enriquecida con Citas y Procedencia")
    st.caption("Descarga el documento con los párrafos marcados resaltados en amarillo y sus notas de origen al pie:")

    en_col1, en_col2 = st.columns(2)
    with en_col1:
        docx_enriched = export_segments(segments, "docx", enriched=True)
        st.download_button(
            label=f"⬇️ Descargar .DOCX Enriquecido ({len(marked_segs)} citas)",
            data=docx_enriched,
            file_name=f"{active_fname.rsplit('.', 1)[0]}_enriquecido_trazabilidad.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
            type="primary" if marked_segs else "secondary",
        )
    with en_col2:
        pdf_enriched = export_segments(segments, "pdf", enriched=True)
        st.download_button(
            label=f"⬇️ Descargar .PDF Enriquecido ({len(marked_segs)} citas)",
            data=pdf_enriched,
            file_name=f"{active_fname.rsplit('.', 1)[0]}_enriquecido_trazabilidad.pdf",
            mime="application/pdf",
            use_container_width=True,
            type="primary" if marked_segs else "secondary",
        )

elif active_context and active_context.get("error"):
    st.error(f"❌ Error al procesar el archivo: {active_context['error']}")
