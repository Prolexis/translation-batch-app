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

import sys
import os

# Garantizar que el directorio raíz del proyecto esté en sys.path (imprescindible en Streamlit Cloud)
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

import io
import re
import html
import time
import base64
import zipfile
import logging
from typing import List, Set, Optional, Dict, Any, Tuple

import streamlit as st

# Configuración de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("translation_app.ui")

from config import settings
try:
    from utils.file_handlers import (
        Segment,
        extract_academic_segments,
        export_segments,
        export_citations_dossier,
        highlight_citations_html,
        UnsupportedFormatError,
        CorruptFileError,
    )
except Exception as _import_exc:
    import traceback
    logger.error("Error crítico importando utils.file_handlers: %s", _import_exc)
    st.error(f"❌ Error al inicializar componentes del sistema (`utils.file_handlers`): {_import_exc}")
    st.code(traceback.format_exc())
    st.stop()

from utils.colors import color_for_index, render_highlighted_html
from agents.orchestrator import TranslationOrchestrator

# Configuración de página en Streamlit
st.set_page_config(
    page_title="Traductor de Papers con Trazabilidad de Origen",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Inicializar selector de tema independiente antes del renderizado
if "theme_choice" not in st.session_state:
    st.session_state["theme_choice"] = "light"

_active_theme = st.session_state.get("theme_choice", "light")

if _active_theme == "dark":
    _theme_vars = """
        --paper-text: #F8FAFC !important;
        --paper-subtext: #94A3B8 !important;
        --paper-card-bg: rgba(30, 41, 59, 0.7) !important;
        --paper-card-border: #334155 !important;
        --paper-sheet-bg: #0B0F19 !important;
        --paper-sheet-border: #1E293B !important;
        --paper-sheet-title: #F8FAFC !important;
        --paper-sheet-subtitle: #94A3B8 !important;
        --paper-sheet-abstract-bg: rgba(30, 41, 59, 0.7) !important;
        --paper-sheet-abstract-border: #334155 !important;
        --paper-sheet-abstract-text: #CBD5E1 !important;
        --paper-sheet-abstract-label: #60A5FA !important;
        --paper-sheet-heading: #93C5FD !important;
        --hero-bg: linear-gradient(135deg, rgba(15, 23, 42, 0.95) 0%, rgba(30, 41, 59, 0.9) 100%) !important;
        --hero-border: rgba(99, 102, 241, 0.4) !important;
        --hero-title-gradient: linear-gradient(135deg, #818CF8 0%, #C084FC 45%, #38BDF8 100%) !important;
        --hero-text: #CBD5E1 !important;
        --hero-pill-bg: rgba(255, 255, 255, 0.08) !important;
        --hero-pill-border: rgba(255, 255, 255, 0.18) !important;
        --hero-pill-text: #E2E8F0 !important;
        --citation-normal-color: #A5B4FC !important;
        --citation-normal-bg: rgba(99, 102, 241, 0.22) !important;
        --citation-normal-border: rgba(99, 102, 241, 0.4) !important;
        --citation-marked-color: #1E1B4B !important;
        --citation-marked-bg: #FDE047 !important;
        --citation-marked-border: #CA8A04 !important;
        --bilingual-normal-bg: rgba(255, 255, 255, 0.04) !important;
        --bilingual-normal-border: #6366F1 !important;
        --bilingual-normal-text: #E2E8F0 !important;
        --marked-card-bg: rgba(254, 240, 138, 0.18) !important;
        --marked-card-border: #F59E0B !important;
        --marked-card-text: #FEF08A !important;
        --marked-badge-bg: #FDE047 !important;
        --marked-badge-text: #1E1B4B !important;
        --marked-foot-text: #FDE047 !important;
    """
    _global_overrides = """
        .stApp {
            background-color: #0E1117 !important;
            color: #F8FAFC !important;
        }
        header[data-testid="stHeader"] {
            background-color: rgba(14, 17, 23, 0.95) !important;
        }
        section[data-testid="stSidebar"] {
            background-color: #161B22 !important;
            border-right: 1px solid #30363D !important;
        }
        section[data-testid="stSidebar"] * {
            color: #F8FAFC !important;
        }
        /* Botones secundarios (☀️ Modo Claro, ❌ Quitar, etc.) en Modo Oscuro */
        button,
        button[kind="secondary"],
        div[data-testid="stButton"] button,
        .stButton button,
        .st-emotion-cache-en1taq,
        .st-emotion-cache-42gn6l,
        button.eqzt73c2 {
            background-color: #21262D !important;
            color: #F8FAFC !important;
            border: 1px solid #30363D !important;
            cursor: pointer !important;
        }
        button *,
        button[kind="secondary"] *,
        div[data-testid="stButton"] button *,
        .stButton button * {
            color: #F8FAFC !important;
            background-color: transparent !important;
        }
        button:hover,
        button[kind="secondary"]:hover,
        div[data-testid="stButton"] button:hover,
        .stButton button:hover {
            background-color: #30363D !important;
            color: #FFFFFF !important;
            border-color: #8B949E !important;
        }
        /* Botones primarios en Modo Oscuro */
        button[kind="primary"],
        button[data-testid="baseButton-primary"],
        button[data-testid="stBaseButton-primary"],
        div[data-testid="stButton"] button[kind="primary"] {
            background-color: #4F46E5 !important;
            color: #FFFFFF !important;
            border: 1px solid #6366F1 !important;
        }
        button[kind="primary"] *,
        button[data-testid="baseButton-primary"] * {
            color: #FFFFFF !important;
        }
        button[kind="primary"]:hover {
            background-color: #4338CA !important;
        }
        /* File Uploader en Modo Oscuro */
        [data-testid="stFileUploader"],
        [data-testid="stFileUploader"] section,
        [data-testid="stFileUploader"] section > div,
        [data-testid="stFileUploaderDropzone"] {
            background-color: #161B22 !important;
            border: 1px dashed #30363D !important;
            color: #F8FAFC !important;
        }
        [data-testid="stFileUploader"] button {
            background-color: #21262D !important;
            color: #F8FAFC !important;
            border: 1px solid #30363D !important;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] span {
            color: #8B949E !important;
        }
        /* Selectbox / dropdowns en Modo Oscuro */
        [data-testid="stSelectbox"],
        [data-testid="stSelectbox"] > div,
        [data-testid="stSelectbox"] div,
        div[data-baseweb="select"],
        div[data-baseweb="select"] > div,
        div[data-baseweb="select"] div {
            background-color: #161B22 !important;
            color: #F8FAFC !important;
            border-color: #30363D !important;
        }
        div[data-baseweb="select"] input,
        [data-testid="stSelectbox"] input {
            background-color: #161B22 !important;
            color: #F8FAFC !important;
            caret-color: #F8FAFC !important;
        }
        div[data-baseweb="select"] button,
        [data-testid="stSelectbox"] button {
            background-color: transparent !important;
            border: none !important;
        }
        div[data-baseweb="select"] svg,
        [data-testid="stSelectbox"] svg {
            fill: #F8FAFC !important;
        }
        div[data-baseweb="popover"],
        div[data-baseweb="popover"] *,
        div[data-baseweb="menu"],
        div[data-baseweb="menu"] *,
        ul[data-baseweb="menu"],
        li[data-baseweb="menu-item"],
        [data-testid="stTooltipContent"],
        div[role="tooltip"] {
            background-color: #161B22 !important;
            color: #F8FAFC !important;
            border-color: #30363D !important;
        }
        li[data-baseweb="menu-item"]:hover,
        li[data-baseweb="menu-item"]:hover * {
            background-color: #21262D !important;
            color: #FFFFFF !important;
        }
        /* Expanders y Contenedores */
        div[data-testid="stExpander"],
        div[data-testid="stExpander"] details,
        div[data-testid="stExpander"] summary {
            background-color: #161B22 !important;
            color: #F8FAFC !important;
            border-color: #334155 !important;
        }
        div[data-testid="stExpander"] summary * {
            color: #F8FAFC !important;
        }
        div[data-testid="stContainer"] {
            border-color: #334155 !important;
        }
        textarea, input,
        input[type="text"], input[type="password"],
        textarea:disabled, input:disabled,
        [data-testid="stTextInput"] input,
        [data-testid="stTextArea"] textarea,
        div[data-baseweb="input"],
        div[data-baseweb="input"] input,
        div[data-baseweb="base-input"],
        div[data-baseweb="textarea"] textarea {
            background-color: #161B22 !important;
            color: #F8FAFC !important;
            -webkit-text-fill-color: #F8FAFC !important;
            opacity: 1 !important;
            border: 1px solid #334155 !important;
        }
        div[data-testid="stMetricValue"] {
            color: #F8FAFC !important;
        }
        div[data-testid="stMetricLabel"] p {
            color: #94A3B8 !important;
        }
        /* Bloques de código (st.code) legibles en Modo Oscuro */
        [data-testid="stCodeBlock"],
        [data-testid="stCodeBlock"] pre,
        [data-testid="stCodeBlock"] code,
        [data-testid="stCodeBlock"] span,
        [data-testid="stCodeBlock"] code *,
        pre code,
        pre code span {
            color: #F0F6FC !important;
            background-color: #0D1117 !important;
        }
        [data-testid="stCodeBlock"] code span,
        [data-testid="stCodeBlock"] code * {
            background: transparent !important;
        }
        p:not([data-testid="stCodeBlock"] *):not(code *):not(pre *),
        span:not([data-testid="stCodeBlock"] *):not(code *):not(pre *),
        label, h1, h2, h3, h4, h5, h6 {
            color: #F8FAFC;
        }
    """
else:
    _theme_vars = """
        --paper-text: #0F172A !important;
        --paper-subtext: #475569 !important;
        --paper-card-bg: #FFFFFF !important;
        --paper-card-border: #CBD5E1 !important;
        --paper-sheet-bg: #FFFFFF !important;
        --paper-sheet-border: #CBD5E1 !important;
        --paper-sheet-title: #0F172A !important;
        --paper-sheet-subtitle: #475569 !important;
        --paper-sheet-abstract-bg: #F8FAFC !important;
        --paper-sheet-abstract-border: #CBD5E1 !important;
        --paper-sheet-abstract-text: #1E293B !important;
        --paper-sheet-abstract-label: #2563EB !important;
        --paper-sheet-heading: #1D4ED8 !important;
        --hero-bg: linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%) !important;
        --hero-border: #818CF8 !important;
        --hero-title-gradient: linear-gradient(135deg, #1E1B4B 0%, #3730A3 50%, #0369A1 100%) !important;
        --hero-text: #1E293B !important;
        --hero-pill-bg: #FFFFFF !important;
        --hero-pill-border: #C7D2FE !important;
        --hero-pill-text: #312E81 !important;
        --citation-normal-color: #3730A3 !important;
        --citation-normal-bg: #EEF2FF !important;
        --citation-normal-border: #818CF8 !important;
        --citation-marked-color: #1E1B4B !important;
        --citation-marked-bg: #FDE047 !important;
        --citation-marked-border: #CA8A04 !important;
        --bilingual-normal-bg: #F8FAFC !important;
        --bilingual-normal-border: #CBD5E1 !important;
        --bilingual-normal-text: #0F172A !important;
        --marked-card-bg: #FEF9C3 !important;
        --marked-card-border: #CA8A04 !important;
        --marked-card-text: #713F12 !important;
        --marked-badge-bg: #FDE047 !important;
        --marked-badge-text: #713F12 !important;
        --marked-foot-text: #854D0E !important;
    """
    _global_overrides = """
        .stApp {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
        }
        header[data-testid="stHeader"] {
            background-color: rgba(255, 255, 255, 0.95) !important;
        }
        section[data-testid="stSidebar"] {
            background-color: #F8FAFC !important;
            border-right: 1px solid #E2E8F0 !important;
        }
        section[data-testid="stSidebar"] * {
            color: #0F172A !important;
        }
        /* Botones secundarios en Modo Claro */
        button,
        button[kind="secondary"],
        div[data-testid="stButton"] button,
        .stButton button,
        .st-emotion-cache-en1taq,
        .st-emotion-cache-42gn6l,
        button.eqzt73c2 {
            background-color: #F1F5F9 !important;
            color: #0F172A !important;
            border: 1px solid #CBD5E1 !important;
            cursor: pointer !important;
        }
        button *,
        button[kind="secondary"] *,
        div[data-testid="stButton"] button *,
        .stButton button * {
            color: #0F172A !important;
            background-color: transparent !important;
        }
        button:hover,
        button[kind="secondary"]:hover,
        div[data-testid="stButton"] button:hover,
        .stButton button:hover {
            background-color: #E2E8F0 !important;
            color: #0F172A !important;
            border-color: #94A3B8 !important;
        }
        /* Botones primarios en Modo Claro */
        button[kind="primary"],
        button[data-testid="baseButton-primary"],
        button[data-testid="stBaseButton-primary"],
        div[data-testid="stButton"] button[kind="primary"] {
            background-color: #4F46E5 !important;
            color: #FFFFFF !important;
            border: none !important;
        }
        button[kind="primary"] *,
        button[data-testid="baseButton-primary"] * {
            color: #FFFFFF !important;
        }
        button[kind="primary"]:hover {
            background-color: #4338CA !important;
        }
        /* File Uploader en Modo Claro */
        [data-testid="stFileUploader"],
        [data-testid="stFileUploader"] section,
        [data-testid="stFileUploader"] section > div,
        [data-testid="stFileUploaderDropzone"] {
            background-color: #F8FAFC !important;
            border: 1px dashed #CBD5E1 !important;
            color: #0F172A !important;
        }
        [data-testid="stFileUploader"] button {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border: 1px solid #CBD5E1 !important;
        }
        [data-testid="stFileUploader"] small,
        [data-testid="stFileUploader"] span {
            color: #475569 !important;
        }
        /* Selectbox / dropdowns en Modo Claro */
        [data-testid="stSelectbox"],
        [data-testid="stSelectbox"] > div,
        [data-testid="stSelectbox"] div,
        div[data-baseweb="select"],
        div[data-baseweb="select"] > div,
        div[data-baseweb="select"] div {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border-color: #CBD5E1 !important;
        }
        div[data-baseweb="select"] input,
        [data-testid="stSelectbox"] input {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            caret-color: #0F172A !important;
        }
        div[data-baseweb="select"] button,
        [data-testid="stSelectbox"] button {
            background-color: transparent !important;
            border: none !important;
        }
        div[data-baseweb="select"] svg,
        [data-testid="stSelectbox"] svg {
            fill: #0F172A !important;
        }
        /* Expanders y Contenedores */
        div[data-testid="stExpander"],
        div[data-testid="stExpander"] details,
        div[data-testid="stExpander"] summary {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            border-color: #CBD5E1 !important;
        }
        div[data-testid="stExpander"] summary * {
            color: #0F172A !important;
        }
        div[data-testid="stContainer"] {
            border-color: #CBD5E1 !important;
        }
        textarea, input,
        input[type="text"], input[type="password"],
        textarea:disabled, input:disabled,
        [data-testid="stTextInput"] input,
        [data-testid="stTextArea"] textarea,
        div[data-baseweb="input"],
        div[data-baseweb="input"] input,
        div[data-baseweb="base-input"],
        div[data-baseweb="textarea"] textarea {
            background-color: #FFFFFF !important;
            color: #0F172A !important;
            -webkit-text-fill-color: #0F172A !important;
            opacity: 1 !important;
            border: 1px solid #CBD5E1 !important;
        }
        div[data-testid="stMetricValue"] {
            color: #0F172A !important;
        }
        div[data-testid="stMetricLabel"] p {
            color: #475569 !important;
        }
        /* Bloques de código (st.code) legibles con alto contraste en Modo Claro */
        [data-testid="stCodeBlock"],
        [data-testid="stCodeBlock"] pre,
        [data-testid="stCodeBlock"] code,
        [data-testid="stCodeBlock"] span,
        [data-testid="stCodeBlock"] code *,
        pre code,
        pre code span {
            color: #F0F6FC !important;
            background-color: #0D1117 !important;
        }
        [data-testid="stCodeBlock"] code span,
        [data-testid="stCodeBlock"] code * {
            background: transparent !important;
        }
        [data-testid="stCodeBlock"] pre {
            border: 1px solid #30363D !important;
            border-radius: 8px !important;
        }
        p:not([data-testid="stCodeBlock"] *):not(code *):not(pre *),
        span:not([data-testid="stCodeBlock"] *):not(code *):not(pre *),
        label, h1, h2, h3, h4, h5, h6 {
            color: #0F172A;
        }
    """

# 0. Protección de React DOM contra mutaciones de Google Translate / extensiones (evita NotFoundError: removeChild)
st.markdown("""
<div style="display:none" aria-hidden="true">
<img src="x" style="display:none" onerror="(function(){
    var targetWin = window;
    try { if (window.parent && window.parent.Node) targetWin = window.parent; } catch(e) {}
    if (targetWin._react_dom_patch_applied) return;
    targetWin._react_dom_patch_applied = true;
    try {
        if (typeof targetWin.Node === 'function' && targetWin.Node.prototype) {
            var origRemoveChild = targetWin.Node.prototype.removeChild;
            targetWin.Node.prototype.removeChild = function(child) {
                if (child && child.parentNode !== this) {
                    return child;
                }
                return origRemoveChild.apply(this, arguments);
            };
            var origInsertBefore = targetWin.Node.prototype.insertBefore;
            targetWin.Node.prototype.insertBefore = function(newNode, refNode) {
                if (refNode && refNode.parentNode !== this) {
                    return newNode;
                }
                return origInsertBefore.apply(this, arguments);
            };
        }
        document.documentElement.setAttribute('translate', 'no');
        document.documentElement.classList.add('notranslate');
        if (document.body) {
            document.body.setAttribute('translate', 'no');
            document.body.classList.add('notranslate');
        }
    } catch(e) {}
})()" />
</div>
""", unsafe_allow_html=True)

# 1. Inyección de variables CSS dinámicas de Tema
st.markdown(f"""
<style>
    :root, .stApp {{
        {_theme_vars}
    }}

    {_global_overrides}
</style>
""", unsafe_allow_html=True)

# 2. Inyección de estilos de componentes y tipografía (CSS estático)
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
        background: var(--hero-bg);
        border: 1px solid var(--hero-border);
        border-radius: 18px;
        padding: 26px 32px;
        margin-bottom: 24px;
        box-shadow: 0 10px 30px 0 rgba(0, 0, 0, 0.08);
    }

    .hero-title {
        font-size: 2.25rem;
        font-weight: 800;
        background: var(--hero-title-gradient);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 8px;
    }

    .hero-desc {
        color: var(--hero-text);
        font-size: 1.05rem;
        line-height: 1.6;
        margin-bottom: 14px;
        font-weight: 500;
    }

    .hero-desc strong {
        color: var(--hero-pill-text);
        font-weight: 700;
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
        background: var(--hero-pill-bg);
        border: 1px solid var(--hero-pill-border);
        border-radius: 30px;
        font-size: 0.82rem;
        font-weight: 600;
        color: var(--hero-pill-text);
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }

    /* Previsualización: Hoja de Paper Científico */
    .paper-sheet-container {
        background: var(--paper-sheet-bg);
        border: 1px solid var(--paper-sheet-border);
        border-radius: 12px;
        padding: 28px 36px;
        box-shadow: 0 15px 35px rgba(0,0,0,0.15);
        max-height: 600px;
        overflow-y: auto;
    }

    .paper-sheet-header {
        text-align: center;
        font-size: 0.75rem;
        letter-spacing: 1.5px;
        color: var(--paper-subtext);
        font-weight: 700;
        text-transform: uppercase;
        margin-bottom: 12px;
        border-bottom: 1px solid var(--paper-sheet-border);
        padding-bottom: 8px;
    }

    .paper-sheet-title {
        text-align: center;
        color: var(--paper-sheet-title);
        font-size: 1.45rem;
        font-weight: 800;
        line-height: 1.35;
        margin: 16px 0 8px 0;
    }

    .paper-sheet-subtitle {
        text-align: center;
        color: var(--paper-sheet-subtitle);
        font-size: 0.85rem;
        margin-bottom: 20px;
    }

    .paper-sheet-abstract {
        background: var(--paper-sheet-abstract-bg);
        border: 1px solid var(--paper-sheet-abstract-border);
        border-radius: 8px;
        padding: 14px 18px;
        margin-bottom: 24px;
    }

    .paper-sheet-abstract-label {
        color: var(--paper-sheet-abstract-label);
        font-size: 0.88rem;
        letter-spacing: 0.5px;
        font-weight: 700;
    }

    .paper-sheet-abstract-text {
        font-style: italic;
        color: var(--paper-sheet-abstract-text);
        font-size: 0.90rem;
        line-height: 1.6;
    }

    .paper-sheet-body-wrap {
        border-top: 1px solid var(--paper-sheet-border);
        padding-top: 16px;
    }

    .paper-sheet-heading {
        color: var(--paper-sheet-heading);
        margin-top: 16px;
        margin-bottom: 6px;
        font-size: 1.05rem;
        text-transform: uppercase;
        font-weight: 700;
    }

    .paper-sheet-ref {
        font-size: 0.78rem;
        color: var(--paper-subtext);
        margin-bottom: 4px;
        padding-left: 14px;
        text-indent: -14px;
        line-height: 1.5;
    }

    .paper-sheet-p {
        color: var(--paper-text);
        font-size: 0.90rem;
        line-height: 1.65;
        margin-bottom: 12px;
    }

    .paper-sheet-p-marked {
        color: var(--marked-card-text);
        background: var(--marked-card-bg);
        border-left: 3px solid var(--marked-card-border);
        padding: 6px 12px;
        border-radius: 4px;
        font-size: 0.90rem;
        line-height: 1.65;
        margin-bottom: 12px;
        font-weight: 500;
    }

    .paper-sheet-more {
        text-align: center;
        color: var(--paper-subtext);
        font-size: 0.82rem;
        margin-top: 16px;
    }

    /* Párrafos en Lienzo de Lectura */
    .preview-p-normal-badge {
        color: var(--paper-subtext);
        font-size: 0.75rem;
        font-weight: 600;
        margin-bottom: 6px;
    }

    .preview-p-marked-badge {
        background: var(--marked-badge-bg);
        color: var(--marked-badge-text);
        font-weight: 800;
        font-size: 0.75rem;
        padding: 3px 8px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 8px;
    }

    .paper-paragraph-normal {
        font-size: 0.98rem;
        line-height: 1.65;
        color: var(--paper-text);
    }

    .paper-paragraph-marked {
        font-size: 1.02rem;
        line-height: 1.65;
        color: var(--marked-card-text);
        font-weight: 500;
    }

    .preview-p-foot-marked {
        margin-top: 8px;
        font-size: 0.82rem;
        color: var(--marked-foot-text);
        font-weight: 600;
        font-style: italic;
    }

    /* Vista Bilingüe Sincronizada */
    .bilingual-card-normal-left {
        background: var(--bilingual-normal-bg);
        color: var(--bilingual-normal-text);
        padding: 10px 14px;
        border-radius: 8px;
        margin-bottom: 10px;
        font-size: 0.92rem;
        line-height: 1.6;
        border: 1px solid var(--paper-sheet-border);
        border-left: 4px solid #6366F1;
    }

    .bilingual-card-normal-right {
        background: var(--bilingual-normal-bg);
        color: var(--bilingual-normal-text);
        padding: 10px 14px;
        border-radius: 8px;
        margin-bottom: 10px;
        font-size: 0.92rem;
        line-height: 1.6;
        border: 1px solid var(--paper-sheet-border);
        border-left: 4px solid #10B981;
    }

    .bilingual-card-marked {
        background: var(--marked-card-bg);
        color: var(--marked-card-text);
        padding: 10px 14px;
        border-radius: 8px;
        margin-bottom: 10px;
        font-size: 0.92rem;
        line-height: 1.6;
        border: 1px solid var(--marked-card-border);
        border-left: 4px solid var(--marked-card-border);
        font-weight: 500;
    }

    .bilingual-prov-label {
        font-size: 0.75rem;
        font-weight: 700;
        opacity: 0.85;
    }

    .inspector-marked-banner {
        background: var(--marked-badge-bg);
        color: var(--marked-badge-text);
        font-weight: 800;
        font-size: 0.80rem;
        padding: 4px 10px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 8px;
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
    st.markdown("### 🌓 Modo de Pantalla")
    curr_th = st.session_state.get("theme_choice", "light")
    btn_th_c1, btn_th_c2 = st.columns(2)
    with btn_th_c1:
        if st.button("☀️ Claro", key="btn_side_theme_light", use_container_width=True, type="primary" if curr_th == "light" else "secondary"):
            if curr_th != "light":
                st.session_state.theme_choice = "light"
                st.rerun()
    with btn_th_c2:
        if st.button("🌙 Oscuro", key="btn_side_theme_dark", use_container_width=True, type="primary" if curr_th == "dark" else "secondary"):
            if curr_th != "dark":
                st.session_state.theme_choice = "dark"
                st.rerun()

    st.markdown("---")
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

    with st.expander("⚙️ Opciones Avanzadas (Opcional)"):
        speed_mode = st.selectbox(
            "Velocidad de Traducción",
            options=["ultra", "balanced", "conservative"],
            index=0,
            format_func=lambda x: {
                "ultra": "⚡ Ultra Rápido (Lotes de 16 párrafos · Recomendado)",
                "balanced": "⚖️ Equilibrado (Lotes de 10 párrafos)",
                "conservative": "🛡️ Seguro / Paso a Paso (Lotes de 6 párrafos)",
            }.get(x, x),
            help="Agrupa múltiples párrafos por llamada para traducir en segundos y evitar pausas de rate-limit.",
        )
        alignment_mode = st.selectbox(
            "Método de Alineación",
            options=["position", "semantic"],
            index=0,
            format_func=lambda x: "Rápida por posición (Recomendada)" if x == "position" else "Semántica (Embeddings)",
            help="Determina el algoritmo de correspondencia entre párrafos originales y traducidos.",
        )

    st.markdown("---")
    st.markdown("#### 🧪 Paper de Prueba")
    st.caption("Prueba el sistema con 1 clic usando un paper de muestra con citas [1] y fórmulas:")

    sample_col1, sample_col2 = st.columns(2)
    sample_dir = os.path.join(os.path.dirname(__file__), "sample_files")

    with sample_col1:
        if st.button("📄 Probar PDF", use_container_width=True):
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
        if st.button("📝 Probar DOCX", use_container_width=True):
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


st.markdown("""
<div class="hero-banner">
    <div class="hero-title">🎓 Traductor de Papers Científicos con Formato Académico</div>
    <div class="hero-desc">
        Traduce tus papers y artículos científicos en inglés (<strong>PDF, Word o TXT</strong>) al español conservando su formato original de <strong>dos columnas, fórmulas y citas bibliográficas</strong> con trazabilidad exacta de procedencia.
    </div>
    <div class="agent-pipeline">
        <span class="agent-pill">1️⃣ Sube tu paper en inglés</span>
        <span class="agent-pill">2️⃣ Revisa la previsualización y descarga en PDF o Word</span>
        <span class="agent-pill">3️⃣ Marca párrafos para tus citas y tesis</span>
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
            speed_map = {
                "ultra": (16, 5),
                "balanced": (10, 4),
                "conservative": (6, 3),
            }
            eff_batch_size, eff_workers = speed_map.get(speed_mode, (16, 5))

            orchestrator = TranslationOrchestrator(
                source_lang=source_lang,
                target_lang=target_lang,
                alignment_mode=alignment_mode,
                api_key=api_key_input,
                batch_size=eff_batch_size,
                max_workers=eff_workers,
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
                f'style="border-radius:10px; border:1px solid var(--paper-sheet-border); box-shadow: 0 10px 30px rgba(0,0,0,0.15); background:var(--paper-sheet-bg);">'
                f'</iframe>'
            )
            st.markdown(pdf_embed_html, unsafe_allow_html=True)
        else:
            # Hoja de Paper Científico (Simulador de Paper Real en HTML/CSS adaptable)
            st.caption("Previsualización estructurada con título, abstract enmarcado y citas resaltadas con color:")

            title_txt = next((s.translated or s.original for s in segments if s.element_type == "title"), segments[0].translated or segments[0].original)
            abstract_txt = " ".join(s.translated or s.original for s in segments if s.element_type == "abstract" or s.section.lower() in ["abstract", "resumen"])

            paper_sheet_html = [
                '<div class="paper-sheet-container">',
                '<div class="paper-sheet-header">REVISTA CIENTÍFICA · TRADUCCIÓN ACADÉMICA IEEE</div>',
                f'<h2 class="paper-sheet-title">{html.escape(title_txt)}</h2>',
                '<div class="paper-sheet-subtitle">Artículo Académico Traducido con Preservación de Citas Bibliográficas y Trazabilidad</div>',
            ]

            if abstract_txt:
                clean_abs = re.sub(r"^(?:abstract|resumen)\s*[\:\—\-\.]*\s*", "", abstract_txt, flags=re.I)
                abs_colored = highlight_citations_html(clean_abs, is_marked=False)
                paper_sheet_html.append(
                    f'<div class="paper-sheet-abstract">'
                    f'<strong class="paper-sheet-abstract-label">RESUMEN — </strong>'
                    f'<span class="paper-sheet-abstract-text">{abs_colored}</span>'
                    f'</div>'
                )

            paper_sheet_html.append('<div class="paper-sheet-body-wrap">')
            for s in segments[:18]:
                if s.element_type in ["title", "abstract"] or s.section.lower() in ["abstract", "resumen"]:
                    continue
                if s.element_type == "heading":
                    paper_sheet_html.append(f'<h4 class="paper-sheet-heading">{html.escape(s.translated or s.original)}</h4>')
                elif s.element_type == "reference":
                    paper_sheet_html.append(f'<div class="paper-sheet-ref">{html.escape(s.translated or s.original)}</div>')
                else:
                    c_txt = highlight_citations_html(s.translated or s.original, is_marked=s.is_marked)
                    p_class = "paper-sheet-p-marked" if s.is_marked else "paper-sheet-p"
                    paper_sheet_html.append(f'<p class="{p_class}">{c_txt}</p>')

            if len(segments) > 18:
                paper_sheet_html.append(f'<div class="paper-sheet-more">... y {len(segments) - 18} párrafos más. Explóralos todos o marca citas en la sección 2 a continuación.</div>')

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

    # Pestañas claras y ordenadas
    tab_preview, tab_inspector, tab_reports, tab_citations = st.tabs([
        "📖 1. Leer y Marcar Párrafos",
        "🔬 2. Inspector Detallado (Bilingüe)",
        "📊 3. Resumen y Estadísticas",
        "📋 4. Mis Citas Seleccionadas",
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
                            header_badge = f'<div class="preview-p-marked-badge">⭐ MARCADO PARA CITACIÓN · {html.escape(s.short_provenance)}</div>'
                            prov_foot = f'<div class="preview-p-foot-marked">📌 {html.escape(s.provenance_label)}</div>'
                            p_html = f'<div class="paper-paragraph-marked">{colored_html}</div>'
                        else:
                            header_badge = f'<div class="preview-p-normal-badge">🏷️ {html.escape(s.short_provenance)}</div>'
                            prov_foot = ""
                            p_html = f'<div class="paper-paragraph-normal">{colored_html}</div>'

                        st.markdown(
                            f'{header_badge}'
                            f'{p_html}'
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
                        b_class = "bilingual-card-marked" if is_m else "bilingual-card-normal-left"
                        star = "⭐ " if is_m else ""
                        colored_orig = highlight_citations_html(s.original, is_marked=is_m)
                        left_html_blocks.append(
                            f'<div class="{b_class}">'
                            f'<strong class="bilingual-prov-label">{star}[{html.escape(s.short_provenance)}]</strong><br>'
                            f'{colored_orig}</div>'
                        )
                    st.markdown("\n".join(left_html_blocks), unsafe_allow_html=True)

            with col_right:
                st.markdown(f"#### 🌐 Traducción Generada ({target_lang.upper()})")
                with st.container(height=520):
                    right_html_blocks = []
                    for s in segments:
                        is_m = (s.id in st.session_state.marked_ids)
                        b_class = "bilingual-card-marked" if is_m else "bilingual-card-normal-right"
                        star = "⭐ " if is_m else ""
                        colored_trans = highlight_citations_html(s.translated or s.original, is_marked=is_m)
                        right_html_blocks.append(
                            f'<div class="{b_class}">'
                            f'<strong class="bilingual-prov-label">{star}[{html.escape(s.short_provenance)}]</strong><br>'
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
                st.markdown('<div class="inspector-marked-banner">⭐ PÁRRAFO MARCADO Y REGISTRADO PARA CITACIÓN / AUDITORÍA</div>', unsafe_allow_html=True)
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
                key=f"orig_area_{sel_seg.id}",
                help="Texto fuente original en inglés extraído del documento.",
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
            st.markdown(
                """<div style="background: rgba(59, 130, 246, 0.12); border: 1px solid rgba(59, 130, 246, 0.35); border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; color: var(--paper-text);" translate="no" class="notranslate">
                    ℹ️ Aún no has marcado ningún párrafo. Puedes marcar párrafos en la pestaña de <strong>1. Leer y Marcar Párrafos</strong> o en el <strong>2. Inspector Detallado</strong> para generar tu ficha de citación.
                </div>""",
                unsafe_allow_html=True,
            )
        else:
            st.markdown(
                f"""<div style="background: rgba(16, 185, 129, 0.15); border: 1px solid rgba(16, 185, 129, 0.4); border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; color: var(--paper-text); font-weight: 500;" translate="no" class="notranslate">
                    ✅ Has seleccionado <strong>{len(marked_segs)} extractos</strong> con trazabilidad de origen.
                </div>""",
                unsafe_allow_html=True,
            )

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
    st.caption("Descarga una versión especial del documento con los párrafos que seleccionaste resaltados en amarillo y con su procedencia al pie:")

    if not marked_segs:
        st.markdown(
            """<div style="background: rgba(59, 130, 246, 0.10); border: 1px solid rgba(59, 130, 246, 0.3); border-radius: 8px; padding: 12px 16px; margin-bottom: 16px; color: var(--paper-text);" translate="no" class="notranslate">
                💡 <strong>Consejo:</strong> Si deseas descargar el paper con citas resaltadas en amarillo y su nota de procedencia exacta, selecciona los párrafos que desees en la pestaña <strong>'1. Leer y Marcar Párrafos'</strong>.
            </div>""",
            unsafe_allow_html=True,
        )

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


if __name__ == "__main__":
    if hasattr(st, "runtime") and not st.runtime.exists():
        import sys
        from streamlit.web import cli as stcli
        sys.argv = ["streamlit", "run", sys.argv[0]]
        sys.exit(stcli.main())
