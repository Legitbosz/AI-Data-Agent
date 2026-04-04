import streamlit as st
import pandas as pd
import os
import shutil
import base64
from agent import analyze_data, analyze_two_datasets, run_analysis_code
from exporters import export_to_excel, export_to_word, export_to_pptx
from cleaner import clean_dataframe, get_data_quality_report
from validator import run_full_validation, validation_summary_text, export_validation_report

# ── Tesseract OCR path (Windows) ─────────────────────────────────────────────
try:
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
except ImportError:
    pass  # OCR won't be available but app still works

st.set_page_config(page_title="AI Data Analysis Agent", layout="wide")

# ── Theme toggle ──────────────────────────────────────────────────────────────
st.session_state.dark_mode = False

# ── CSS — full platform styling ───────────────────────────────────────────────
def get_css(dark):
    # ── Colour tokens ──────────────────────────────────────────────────────────
    if dark:
        bg          = "#070d1b"
        bg_card     = "#0e1a2e"
        bg_input    = "#0b1525"
        text        = "#e8edf5"
        text_muted  = "#5a7396"
        border      = "#152038"
        accent      = "#2563eb"
        accent2     = "#4f46e5"
        grid        = "rgba(37,99,235,0.07)"
        glow        = "rgba(37,99,235,0.18)"
        upload_bg   = "#0a1422"
        upload_text = "#8fb3d8"
        btn_browse  = "#1d4ed8"
        shadow      = "0 4px 24px rgba(0,0,0,0.5)"
    else:
        bg          = "#f0f5ff"
        bg_card     = "#ffffff"
        bg_input    = "#ffffff"
        text        = "#0f172a"
        text_muted  = "#64748b"
        border      = "#c7d8f0"
        accent      = "#2563eb"
        accent2     = "#4f46e5"
        grid        = "rgba(37,99,235,0.08)"
        glow        = "rgba(37,99,235,0.12)"
        upload_bg   = "#ffffff"
        upload_text = "#334155"
        btn_browse  = "#2563eb"
        shadow      = "0 4px 24px rgba(37,99,235,0.1)"

    return f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@300;400;600;700;800&display=swap');

/* ═══════════════════════════════════════
   BASE
═══════════════════════════════════════ */
html, body, [class*="css"] {{
    font-family: 'Sora', sans-serif !important;
}}

/* ═══════════════════════════════════════
   APP BACKGROUND
═══════════════════════════════════════ */
.stApp {{
    background: {bg} !important;
    background-image:
        radial-gradient(ellipse 70% 40% at 50% 0%, {glow} 0%, transparent 65%),
        linear-gradient({grid} 1px, transparent 1px),
        linear-gradient(90deg, {grid} 1px, transparent 1px) !important;
    background-size: 100%, 50px 50px, 50px 50px !important;
    min-height: 100vh !important;
}}

.main .block-container {{
    background: transparent !important;
    padding: 0 2rem 2rem !important;
    padding-top: 0 !important;
    margin-top: 0 !important;
}}
.main {{
    padding-top: 0 !important;
    margin-top: 0 !important;
}}
header[data-testid="stHeader"] {{
    display: none !important;
}}
[data-testid="stToolbar"] {{ display: none !important; }}
[data-testid="stDecoration"] {{ display: none !important; }}
[data-testid="stAppViewContainer"] > section {{
    padding-top: 0 !important;
}}

/* Hide sidebar and ALL collapse buttons */
section[data-testid="stSidebar"],
[data-testid="stSidebarCollapseButton"],
[data-testid="collapsedControl"],
[data-testid="stSidebarUserContent"],
div[data-testid="stSidebarCollapsedControl"],
.st-emotion-cache-1q2d14y,
.st-emotion-cache-17lntkn,
.st-emotion-cache-1cypcdb,
.st-emotion-cache-h5rgaw,
button[aria-label="Close sidebar"],
button[aria-label="Open sidebar"] {{
    display: none !important;
    visibility: hidden !important;
    width: 0 !important;
    height: 0 !important;
    overflow: hidden !important;
    position: absolute !important;
    pointer-events: none !important;
}}

/* ═══════════════════════════════════════
   TYPOGRAPHY
═══════════════════════════════════════ */
h1, h2, h3, h4, h5, p, span, div, label,
.stMarkdown, .stText {{
    color: {text} !important;
    font-family: 'Sora', sans-serif !important;
}}
h1 {{ font-size: 1.6rem !important; font-weight: 700 !important; }}
h2 {{ font-size: 1.3rem !important; font-weight: 600 !important; }}
h3 {{ font-size: 1.1rem !important; font-weight: 600 !important; }}

/* ═══════════════════════════════════════
   HEADER BANNER — full-width, touches top
═══════════════════════════════════════ */
.agent-header-wrap {{
    width: calc(100% + 4rem);
    margin-left: -2rem;
    margin-right: -2rem;
    margin-top: -5rem;
    background: linear-gradient(125deg, #1a40c8 0%, {accent2} 45%, #7c3aed 100%);
    box-shadow: 0 4px 30px rgba(37,99,235,0.5);
    overflow: hidden;
    margin-bottom: 20px;
}}
.agent-header-inner {{
    padding: 14px 2rem 12px 2rem;
    position: relative;
    z-index: 1;
}}
.agent-header-wrap::before {{
    content: '';
    position: absolute; top: -50%; right: -5%;
    width: 500px; height: 500px;
    background: radial-gradient(circle, rgba(255,255,255,0.07) 0%, transparent 65%);
    border-radius: 50%; pointer-events: none;
}}
.agent-header-wrap::after {{
    content: '';
    position: absolute; bottom: -60%; left: 30%;
    width: 320px; height: 320px;
    background: radial-gradient(circle, rgba(255,255,255,0.04) 0%, transparent 70%);
    border-radius: 50%; pointer-events: none;
}}
.agent-header {{
    max-width: 900px;
    padding: 36px 0 32px;
    position: relative; z-index: 1;
}}
.agent-title {{
    font-size: 2rem !important;
    font-weight: 800 !important;
    color: #ffffff !important;
    margin: 0 !important;
    letter-spacing: -1.5px !important;
    line-height: 1.1 !important;
    text-shadow: 0 3px 20px rgba(0,0,0,0.3) !important;
}}
.agent-subtitle {{
    font-size: 0.9rem !important;
    color: rgba(255,255,255,0.80) !important;
    margin: 12px 0 18px !important;
    font-weight: 300 !important;
    letter-spacing: 0.3px !important;
}}
.agent-badge {{
    display: inline-flex;
    align-items: center;
    gap: 6px;
    background: rgba(255,255,255,0.13);
    border: 1px solid rgba(255,255,255,0.26);
    border-radius: 24px;
    padding: 5px 16px;
    font-size: 0.78rem;
    font-weight: 600;
    color: white !important;
    letter-spacing: 0.4px;
}}

/* ═══════════════════════════════════════
   THEME TOGGLE — fixed top-right
═══════════════════════════════════════ */
/* Theme toggle — fixed far right, vertically centered in banner */
/* Theme toggle styling */
.stButton button {{
    transition: all 0.2s !important;
}}

/* sidebar removed */

/* ═══════════════════════════════════════
   FILE UPLOADER
═══════════════════════════════════════ */
[data-testid="stFileUploader"] {{
    background: {upload_bg} !important;
    border: 2px dashed {border} !important;
    border-radius: 16px !important;
    padding: 8px !important;
    transition: border-color 0.25s !important;
}}
[data-testid="stFileUploader"]:hover {{
    border-color: {accent} !important;
}}
[data-testid="stFileUploaderDropzone"] {{
    background: {upload_bg} !important;
    border: none !important;
}}
[data-testid="stFileUploaderDropzone"] * {{
    color: {upload_text} !important;
}}
[data-testid="stFileUploaderDropzone"] button {{
    background: {btn_browse} !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: 'Sora', sans-serif !important;
    font-weight: 600 !important;
}}
[data-testid="stFileUploader"] small {{
    color: {text_muted} !important;
}}

/* ═══════════════════════════════════════
   CHAT MESSAGES
═══════════════════════════════════════ */
[data-testid="stChatMessage"] {{
    background: {bg_card} !important;
    border: 1px solid {border} !important;
    border-radius: 16px !important;
    margin-bottom: 14px !important;
    padding: 20px !important;
    box-shadow: {shadow} !important;
}}
[data-testid="stChatMessage"] * {{
    color: {text} !important;
}}
/* Fix chat avatar icon text showing as raw text */
[data-testid="stChatMessageAvatarUser"] {{
    background: {accent} !important;
    border-radius: 50% !important;
    width: 36px !important; height: 36px !important;
    display: flex !important; align-items: center !important;
    justify-content: center !important;
    font-size: 1.1rem !important;
    overflow: hidden !important;
}}
[data-testid="stChatMessageAvatarUser"] span {{
    display: none !important;
}}
[data-testid="stChatMessageAvatarUser"]::after {{
    content: "👤";
    font-size: 1.1rem;
}}
[data-testid="stChatMessageAvatarAssistant"] {{
    background: linear-gradient(135deg, {accent}, {accent2}) !important;
    border-radius: 50% !important;
    width: 36px !important; height: 36px !important;
    display: flex !important; align-items: center !important;
    justify-content: center !important;
    overflow: hidden !important;
}}
[data-testid="stChatMessageAvatarAssistant"] span {{
    display: none !important;
}}
[data-testid="stChatMessageAvatarAssistant"]::after {{
    content: "🤖";
    font-size: 1.1rem;
}}

/* ═══════════════════════════════════════
   CHAT INPUT
═══════════════════════════════════════ */
[data-testid="stBottom"],
[data-testid="stBottom"] > div,
[data-testid="stBottom"] > div > div {{
    background: {bg} !important;
    border-top: 1px solid {border} !important;
    padding: 12px 8% !important;
}}
[data-testid="stChatInput"] {{
    background: {bg_input} !important;
    border: 1.5px solid {border} !important;
    border-radius: 14px !important;
    box-shadow: {shadow} !important;
}}
[data-testid="stChatInput"] > div {{
    background: {bg_input} !important;
}}
[data-testid="stChatInput"] textarea {{
    background: {bg_input} !important;
    color: {text} !important;
    font-family: 'Sora', sans-serif !important;
    font-size: 0.95rem !important;
}}
[data-testid="stChatInput"] textarea::placeholder {{
    color: {text_muted} !important;
    opacity: 1 !important;
}}
[data-testid="stChatInput"]:focus-within {{
    border-color: {accent} !important;
}}
/* Force chat input wrapper background — but not caret */
[data-testid="stChatInput"] > div,
[data-testid="stChatInput"] > div > div {{
    background: {bg_input} !important;
}}
[data-testid="stChatInput"] textarea {{
    caret-color: {accent} !important;
    cursor: text !important;
}}
[data-testid="stChatInput"] button {{
    background: {accent} !important;
    color: white !important;
}}

/* ═══════════════════════════════════════
   METRICS
═══════════════════════════════════════ */
[data-testid="stMetric"] {{
    background: {bg_card} !important;
    border: 1px solid {border} !important;
    border-radius: 12px !important;
    padding: 14px 16px !important;
    box-shadow: {shadow} !important;
}}
[data-testid="stMetricValue"] {{
    color: {accent} !important;
    font-weight: 700 !important;
}}

/* ═══════════════════════════════════════
   DATAFRAMES
═══════════════════════════════════════ */
[data-testid="stDataFrame"] {{
    border: 1px solid {border} !important;
    border-radius: 12px !important;
    overflow: hidden !important;
    box-shadow: {shadow} !important;
}}

/* ═══════════════════════════════════════
   BUTTONS
═══════════════════════════════════════ */
.stButton > button {{
    background: {bg_card} !important;
    color: {text} !important;
    border: 1px solid {border} !important;
    border-radius: 10px !important;
    font-family: 'Sora', sans-serif !important;
    font-weight: 500 !important;
    transition: all 0.2s !important;
}}
.stButton > button:hover {{
    border-color: {accent} !important;
    color: {accent} !important;
    box-shadow: 0 0 0 3px rgba(37,99,235,0.15) !important;
}}

/* ═══════════════════════════════════════
   EXPANDERS
═══════════════════════════════════════ */
[data-testid="stExpander"] {{
    background: {bg_card} !important;
    border: 1px solid {border} !important;
    border-radius: 12px !important;
    box-shadow: {shadow} !important;
}}

/* ═══════════════════════════════════════
   ALERTS / INFO
═══════════════════════════════════════ */
[data-testid="stAlert"] {{
    background: {bg_card} !important;
    border: 1px solid {border} !important;
    border-radius: 12px !important;
    color: {text} !important;
}}

/* ═══════════════════════════════════════
   SELECTBOX / MULTISELECT
═══════════════════════════════════════ */
div[data-baseweb="select"] * {{ cursor: default !important; color: {text} !important; }}
div[data-baseweb="select"] > div {{
    background: {bg_input} !important;
    border-color: {border} !important;
    border-radius: 10px !important;
}}

/* ═══════════════════════════════════════
   DIVIDER
═══════════════════════════════════════ */
hr {{ border-color: {border} !important; opacity: 0.6 !important; }}

/* ═══════════════════════════════════════
   EXPORT SECTION
═══════════════════════════════════════ */
.export-section {{
    background: {bg_card};
    border: 1px solid {border};
    border-radius: 20px;
    padding: 28px 32px;
    margin-top: 24px;
    box-shadow: {shadow};
}}
.export-title {{
    font-size: 0.72rem !important;
    font-weight: 700 !important;
    color: {text_muted} !important;
    text-transform: uppercase !important;
    letter-spacing: 2px !important;
    margin-bottom: 20px !important;
}}

/* ═══════════════════════════════════════
   EXPORT DOWNLOAD BUTTONS (HTML anchors)
═══════════════════════════════════════ */
.export-dl-btn {{
    display: block;
    text-align: center;
    text-decoration: none !important;
    color: white !important;
    font-family: 'Sora', sans-serif;
    font-weight: 600;
    font-size: 0.8rem;
    padding: 12px 18px;
    border-radius: 12px;
    transition: transform 0.2s, box-shadow 0.2s, filter 0.2s;
    cursor: pointer;
    letter-spacing: 0.3px;
}}
.export-dl-btn:hover {{
    transform: translateY(-2px);
    filter: brightness(1.1);
    text-decoration: none !important;
    color: white !important;
}}
.export-dl-excel {{
    background: linear-gradient(135deg, #14532d 0%, #16a34a 100%);
    box-shadow: 0 4px 14px rgba(22,163,74,0.35);
}}
.export-dl-excel:hover {{ box-shadow: 0 8px 24px rgba(22,163,74,0.5); }}
.export-dl-word {{
    background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
    box-shadow: 0 4px 14px rgba(37,99,235,0.35);
}}
.export-dl-word:hover {{ box-shadow: 0 8px 24px rgba(37,99,235,0.5); }}
.export-dl-pptx {{
    background: linear-gradient(135deg, #7c2d12 0%, #f97316 100%);
    box-shadow: 0 4px 14px rgba(234,88,12,0.35);
}}
.export-dl-pptx:hover {{ box-shadow: 0 8px 24px rgba(234,88,12,0.5); }}

/* ═══════════════════════════════════════
   SCROLLBAR
═══════════════════════════════════════ */
::-webkit-scrollbar {{ width: 5px; height: 5px; }}
::-webkit-scrollbar-track {{ background: {bg}; }}
::-webkit-scrollbar-thumb {{ background: {border}; border-radius: 4px; }}
::-webkit-scrollbar-thumb:hover {{ background: {accent}; }}

/* ═══════════════════════════════════════
   SUCCESS / WARNING
═══════════════════════════════════════ */
[data-testid="stSuccess"] {{ background: rgba(22,163,74,0.1) !important; border-radius: 10px !important; }}
[data-testid="stWarning"] {{ background: rgba(234,179,8,0.1) !important; border-radius: 10px !important; }}

/* ═══════════════════════════════════════
   IMAGE FULLSCREEN BUTTON & OVERLAY
═══════════════════════════════════════ */
/* Fullscreen button — target every possible selector */
button[title="Fullscreen"],
button[title="Close fullscreen"],
button[aria-label="Fullscreen"],
button[aria-label="Close fullscreen"],
[data-testid="StyledFullScreenButton"],
[data-testid="stImage"] button,
[data-testid="stImage"] > div > button,
div[class*="fullscreen"] button,
div[class*="FullScreen"] button {{
    background: {bg_card} !important;
    background-color: {bg_card} !important;
    border: 1px solid {border} !important;
    border-radius: 8px !important;
    color: {text} !important;
    opacity: 0.9 !important;
    box-shadow: 0 2px 10px rgba(0,0,0,0.08) !important;
}}
button[title="Fullscreen"] svg,
button[title="Close fullscreen"] svg,
button[aria-label="Fullscreen"] svg,
button[aria-label="Close fullscreen"] svg,
[data-testid="StyledFullScreenButton"] svg,
[data-testid="stImage"] button svg {{
    fill: {text} !important;
    stroke: {text} !important;
    color: {text} !important;
}}

/* Fullscreen overlay — nuclear approach: any fixed overlay with high z-index */
div[data-baseweb="modal"],
div[data-baseweb="modal"] > div,
div[data-baseweb="modal"] > div > div,
div[role="dialog"],
div[role="dialog"] > div,
[data-testid="stModal"],
[data-testid="stModal"] > div {{
    background: {bg} !important;
    background-color: {bg} !important;
}}
/* ═══════════════════════════════════════
   QUICK ACTION CHIPS
═══════════════════════════════════════ */
div[data-testid="stHorizontalBlock"] button[kind="secondary"] {{
    border: 1.5px solid {border} !important;
    border-radius: 20px !important;
    background: {bg_card} !important;
    color: {text_muted} !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    padding: 8px 12px !important;
    transition: all 0.2s ease !important;
    white-space: nowrap !important;
}}
div[data-testid="stHorizontalBlock"] button[kind="secondary"]:hover {{
    border-color: {accent} !important;
    color: {accent} !important;
    background: {bg_input} !important;
    box-shadow: 0 2px 12px rgba(37,99,235,0.15) !important;
    transform: translateY(-1px) !important;
}}

</style>"""

st.markdown(get_css(st.session_state.dark_mode), unsafe_allow_html=True)

# ── JS fix: inject <style> into parent document head to override Streamlit emotion-cache ──
import streamlit.components.v1 as components

_dark = st.session_state.dark_mode
_light_bg = "#f0f5ff" if not _dark else "#070d1b"
_card_bg  = "#ffffff" if not _dark else "#0e1a2e"
_text_clr = "#0f172a" if not _dark else "#e8edf5"
_border   = "#c7d8f0" if not _dark else "#152038"

components.html(f"""
<script>
(function() {{
    var doc = window.parent.document;
    // Remove any previous injection
    var old = doc.getElementById('custom-fullscreen-fix');
    if (old) old.remove();

    var style = doc.createElement('style');
    style.id = 'custom-fullscreen-fix';
    style.textContent = `
        /* ── FULLSCREEN BUTTON (the small icon on images) ── */
        [data-testid="StyledFullScreenButton"],
        [data-testid="StyledFullScreenButton"] > button,
        [data-testid="stImage"] button,
        [data-testid="stFullScreenFrame"] button,
        button[title="Fullscreen"],
        button[title="Close fullscreen"],
        button[kind="minimal"][class*="emotion"],
        div[class*="emotion"] > button[kind="minimal"] {{
            background: {_card_bg} !important;
            background-color: {_card_bg} !important;
            border: 1.5px solid {_border} !important;
            border-radius: 10px !important;
            color: {_text_clr} !important;
            opacity: 1 !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08) !important;
            padding: 6px !important;
        }}
        [data-testid="StyledFullScreenButton"] svg,
        [data-testid="StyledFullScreenButton"] svg *,
        [data-testid="stImage"] button svg,
        [data-testid="stImage"] button svg *,
        button[title="Fullscreen"] svg,
        button[title="Fullscreen"] svg *,
        button[title="Close fullscreen"] svg,
        button[title="Close fullscreen"] svg * {{
            fill: {_text_clr} !important;
            stroke: {_text_clr} !important;
            color: {_text_clr} !important;
        }}

        /* ── FULLSCREEN OVERLAY (the dark backdrop when viewing image) ── */
        div[data-baseweb="modal"],
        div[data-baseweb="modal"] > div,
        div[data-baseweb="modal"] > div > div,
        div[data-baseweb="modal"] > div > div > div,
        [role="dialog"],
        [role="dialog"] > div {{
            background: {_light_bg} !important;
            background-color: {_light_bg} !important;
        }}

        /* ── CLOSE BUTTON inside fullscreen overlay ── */
        div[data-baseweb="modal"] button,
        [role="dialog"] button,
        button[aria-label="Close"] {{
            background: {_card_bg} !important;
            background-color: {_card_bg} !important;
            border: 1.5px solid {_border} !important;
            border-radius: 10px !important;
            color: {_text_clr} !important;
        }}
        div[data-baseweb="modal"] button svg,
        div[data-baseweb="modal"] button svg *,
        [role="dialog"] button svg,
        [role="dialog"] button svg *,
        button[aria-label="Close"] svg,
        button[aria-label="Close"] svg * {{
            fill: {_text_clr} !important;
            stroke: {_text_clr} !important;
            color: {_text_clr} !important;
        }}

        /* ── Tooltip fix ── */
        div[role="tooltip"],
        div[data-baseweb="tooltip"],
        div[data-baseweb="tooltip"] > div {{
            background: {_card_bg} !important;
            background-color: {_card_bg} !important;
            color: {_text_clr} !important;
            border: 1px solid {_border} !important;
            border-radius: 8px !important;
        }}
    `;
    doc.head.appendChild(style);
}})();
</script>
""", height=0)


# ── Header — full width banner via CSS injection ───────────────────────────────
st.markdown(f"""
<div class="agent-header-wrap">
    <div class="agent-header-inner">
        <div>
            <p class="agent-title">🤖 AI Data Analysis Agent</p>
            <p class="agent-subtitle">Upload CSV, Excel & PDF &nbsp;·&nbsp; Ask in plain English &nbsp;·&nbsp; Export to Excel, Word & PowerPoint</p>
            <span class="agent-badge">⚡ Powered by Claude AI (Anthropic)</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Theme toggle — fixed top-right corner of the screen

for key, default in {
    "messages": [], "df1": None, "df2": None,
    "name1": "Dataset 1", "name2": "Dataset 2",
    "last_analysis": None, "validation_results": None
}.items():
    if key not in st.session_state:
        st.session_state[key] = default


def load_file(uploaded, header_row=0):
    if uploaded.name.endswith(".csv"):
        df = pd.read_csv(uploaded, header=header_row)
    elif uploaded.name.endswith(".pdf"):
        dfs = load_pdf_tables(uploaded)
        if not dfs:
            raise ValueError("No tables found in this PDF.")
        df = dfs[0]
    else:
        df = pd.read_excel(uploaded, header=header_row)
    df.columns = df.columns.astype(str).str.replace(r'[\s]+', ' ', regex=True).str.strip()
    return df


def _clean_pdf_df(df):
    """Clean a DataFrame extracted from PDF: fix columns, convert numerics."""
    df = df.dropna(how='all').dropna(axis=1, how='all')
    # Drop columns that are entirely empty strings
    df = df.loc[:, ~(df.astype(str).eq('').all())]
    for col in df.columns:
        cleaned = df[col].astype(str).str.replace(r'[\u20a6$\u20ac\u00a3,\s]', '', regex=True)
        try:
            numeric = pd.to_numeric(cleaned, errors='coerce')
            if numeric.notna().sum() > len(df) * 0.5:
                df[col] = numeric
        except Exception:
            pass
    return df


def _ocr_pdf_to_tables(uploaded):
    """OCR fallback: convert PDF pages to images, run Tesseract, parse tables."""
    import logging
    tables = []

    try:
        from pdf2image import convert_from_bytes
        import pytesseract
    except ImportError as e:
        logging.warning(f"[PDF-OCR] Missing library: {e}. Install pdf2image and pytesseract.")
        return tables

    uploaded.seek(0)
    pdf_bytes = uploaded.read()
    uploaded.seek(0)

    try:
        images = convert_from_bytes(pdf_bytes, dpi=300)
    except Exception as e:
        logging.warning(f"[PDF-OCR] pdf2image failed: {e}")
        return tables

    for page_num, img in enumerate(images, start=1):
        try:
            # Try to get structured table data via Tesseract TSV output
            tsv_text = pytesseract.image_to_data(img, output_type=pytesseract.Output.DATAFRAME)
            # Filter to meaningful text
            tsv_text = tsv_text[tsv_text['text'].notna() & (tsv_text['text'].astype(str).str.strip() != '')]

            if len(tsv_text) < 3:
                continue

            # Group by block/line to reconstruct rows
            rows = []
            for (block, par, line), group in tsv_text.groupby(['block_num', 'par_num', 'line_num']):
                row_text = group.sort_values('left')['text'].astype(str).tolist()
                if row_text:
                    rows.append(row_text)

            if len(rows) < 2:
                continue

            # Normalize: find the most common row length (likely table width)
            from collections import Counter
            lengths = Counter(len(r) for r in rows)
            most_common_len = lengths.most_common(1)[0][0]

            # Filter rows that match the most common length (likely table rows)
            table_rows = [r for r in rows if len(r) == most_common_len]

            if len(table_rows) < 2:
                # Fall back: pad/trim all rows to most common length
                table_rows = []
                for r in rows:
                    if len(r) >= most_common_len:
                        table_rows.append(r[:most_common_len])
                    elif len(r) >= most_common_len - 1:
                        table_rows.append(r + [''] * (most_common_len - len(r)))

            if len(table_rows) < 2:
                continue

            header = [str(c).strip() if c else f"Col_{i}" for i, c in enumerate(table_rows[0])]
            df = pd.DataFrame(table_rows[1:], columns=header)
            df = _clean_pdf_df(df)

            if len(df) > 0 and len(df.columns) > 0:
                df.attrs['_pdf_source'] = f"Page {page_num} (OCR)"
                tables.append(df)

        except Exception as e:
            logging.warning(f"[PDF-OCR] Page {page_num} failed: {e}")
            continue

    # If structured approach failed, try plain text OCR
    if not tables:
        try:
            all_text = []
            for img in images:
                text = pytesseract.image_to_string(img)
                if text and text.strip():
                    all_text.append(text.strip())

            if all_text:
                from io import StringIO
                combined = '\n'.join(all_text)
                lines = [l for l in combined.split('\n') if l.strip()]
                if len(lines) >= 2:
                    try:
                        df = pd.read_csv(StringIO('\n'.join(lines)), sep=r'\s{2,}', engine='python')
                        if len(df) > 0 and len(df.columns) > 1:
                            df = _clean_pdf_df(df)
                            df.attrs['_pdf_source'] = "Full OCR text"
                            tables.append(df)
                    except Exception:
                        pass
        except Exception as e:
            logging.warning(f"[PDF-OCR] Plain text fallback failed: {e}")

    return tables


def load_pdf_tables(uploaded):
    """
    Extract tables from PDF using a 3-tier approach:
    1. pdfplumber (fast, for native/digital PDFs)
    2. pdfplumber text extraction (for text-only PDFs without table structure)
    3. OCR via pytesseract (for scanned/image PDFs)
    """
    import pdfplumber
    import logging

    uploaded.seek(0)
    all_tables = []

    # ── Tier 1: pdfplumber table extraction ──
    with pdfplumber.open(uploaded) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables()
            for t_idx, table in enumerate(tables):
                if not table or len(table) < 2:
                    continue
                header = [str(c).strip() if c else f"Col_{i}" for i, c in enumerate(table[0])]
                rows = table[1:]
                df = pd.DataFrame(rows, columns=header)
                df = _clean_pdf_df(df)
                if len(df) > 0 and len(df.columns) > 0:
                    df.attrs['_pdf_source'] = f"Page {page_num}, Table {t_idx + 1}"
                    all_tables.append(df)

    if all_tables:
        uploaded.seek(0)
        return all_tables

    # ── Tier 2: pdfplumber text extraction ──
    uploaded.seek(0)
    with pdfplumber.open(uploaded) as pdf:
        lines_txt = []
        has_text = False
        for page in pdf.pages:
            text = page.extract_text()
            if text and text.strip():
                has_text = True
                lines_txt.extend(text.strip().split('\n'))

        if lines_txt:
            from io import StringIO
            text_block = '\n'.join(lines_txt)
            try:
                df = pd.read_csv(StringIO(text_block), sep=r'\s{2,}', engine='python')
                if len(df) > 0 and len(df.columns) > 1:
                    df = _clean_pdf_df(df)
                    df.attrs['_pdf_source'] = "Text extraction"
                    all_tables.append(df)
            except Exception:
                pass

    if all_tables:
        uploaded.seek(0)
        return all_tables

    # ── Tier 3: OCR fallback (for scanned/image PDFs) ──
    logging.info("[PDF] No tables found via text extraction, trying OCR...")
    uploaded.seek(0)
    ocr_tables = _ocr_pdf_to_tables(uploaded)
    all_tables.extend(ocr_tables)

    uploaded.seek(0)
    return all_tables




def preview_raw(uploaded, n=10):
    try:
        uploaded.seek(0)
        if uploaded.name.endswith(".pdf"):
            return None  # PDF preview handled separately
        elif uploaded.name.endswith(".csv"):
            return pd.read_csv(uploaded, header=None, nrows=n)
        else:
            return pd.read_excel(uploaded, header=None, nrows=n)
    except Exception:
        return None
    finally:
        uploaded.seek(0)


def _handle_pdf_upload(uploaded, slot_prefix, col_widget):
    """Handle PDF upload with multi-table support. Returns True if loaded."""
    uploaded.seek(0)
    with col_widget.status("Extracting data from PDF...", expanded=True) as status:
        st.write("Trying direct table extraction...")
        pdf_tables = load_pdf_tables(uploaded)
        uploaded.seek(0)
        if pdf_tables:
            src = pdf_tables[0].attrs.get('_pdf_source', '')
            if 'OCR' in src:
                status.update(label="Extracted via OCR (scanned PDF)", state="complete")
            elif 'Text' in src:
                status.update(label="Extracted from text content", state="complete")
            else:
                status.update(label="Extracted structured tables", state="complete")
        else:
            status.update(label="No data found", state="error")

    if not pdf_tables:
        col_widget.warning(
            "No tables found in this PDF. If it's a scanned/image PDF, "
            "make sure **Tesseract OCR** and **Poppler** are installed on your system."
        )
        return False

    base_name = uploaded.name.rsplit(".", 1)[0]

    if len(pdf_tables) == 1:
        df = pdf_tables[0]
        st.session_state[f"df{slot_prefix}"] = df
        st.session_state[f"name{slot_prefix}"] = base_name
        col_widget.success(f"Loaded: {df.shape[0]} rows x {df.shape[1]} cols")
        col_widget.dataframe(df.head(5), width='stretch')
        return True
    else:
        col_widget.info(f"Found **{len(pdf_tables)} tables** in PDF")

        # Show all tables with selection
        table_labels = []
        for i, df in enumerate(pdf_tables):
            src = df.attrs.get('_pdf_source', f'Table {i+1}')
            table_labels.append(f"{src} ({df.shape[0]} rows x {df.shape[1]} cols)")

        selected = col_widget.selectbox(
            "Select table to analyze:",
            range(len(pdf_tables)),
            format_func=lambda i: table_labels[i],
            key=f"pdf_table_select_{slot_prefix}"
        )

        df = pdf_tables[selected]
        st.session_state[f"df{slot_prefix}"] = df
        st.session_state[f"name{slot_prefix}"] = f"{base_name} - {table_labels[selected].split(' (')[0]}"
        col_widget.success(f"Loaded: {df.shape[0]} rows x {df.shape[1]} cols")
        col_widget.dataframe(df.head(5), width='stretch')

        # Also store all tables so user can switch later
        st.session_state[f"_pdf_tables_{slot_prefix}"] = pdf_tables

        # Show previews of other tables in an expander
        with col_widget.expander(f"Preview all {len(pdf_tables)} tables"):
            for i, tdf in enumerate(pdf_tables):
                src = tdf.attrs.get('_pdf_source', f'Table {i+1}')
                st.markdown(f"**{src}** ({tdf.shape[0]} rows x {tdf.shape[1]} cols)")
                st.dataframe(tdf.head(3), width='stretch')
                st.divider()
        return True


# ── MAIN AREA: uploaders run first so session_state is populated ───────────────
st.subheader("Upload Files")
col1, col2 = st.columns(2)

with col1:
    st.markdown("**File 1**")
    file1 = st.file_uploader("Upload first file", type=["csv","xlsx","xls","pdf"], key="file1")
    if file1:
        if file1.name.endswith(".pdf"):
            _handle_pdf_upload(file1, "1", col1)
        else:
            raw1 = preview_raw(file1)
            st.session_state._raw1      = raw1
            st.session_state._file1_obj = file1
            if raw1 is not None:
                st.markdown("**Preview (raw):**")
                st.dataframe(raw1, width='stretch')
            file1.seek(0)
            header1 = int(st.session_state.get("header1_sb", 0) or 0) if isinstance(st.session_state.get("header1_sb", 0), int) else 0
            st.session_state.df1   = load_file(file1, header1)
            st.session_state.name1 = file1.name.rsplit(".", 1)[0]
            st.success(f"Loaded: {st.session_state.df1.shape[0]} rows x {st.session_state.df1.shape[1]} cols")
            st.dataframe(st.session_state.df1.head(5), width='stretch')

with col2:
    st.markdown("**File 2 (optional \u2014 for comparison)**")
    file2 = st.file_uploader("Upload second file", type=["csv","xlsx","xls","pdf"], key="file2")
    if file2:
        if file2.name.endswith(".pdf"):
            _handle_pdf_upload(file2, "2", col2)
        else:
            raw2 = preview_raw(file2)
            st.session_state._raw2      = raw2
            st.session_state._file2_obj = file2
            if raw2 is not None:
                st.markdown("**Preview (raw):**")
                st.dataframe(raw2, width='stretch')
            file2.seek(0)
            header2 = int(st.session_state.get("header2_sb", 0) or 0) if isinstance(st.session_state.get("header2_sb", 0), int) else 0
            st.session_state.df2   = load_file(file2, header2)
            st.session_state.name2 = file2.name.rsplit(".", 1)[0]
            st.success(f"Loaded: {st.session_state.df2.shape[0]} rows x {st.session_state.df2.shape[1]} cols")
            st.dataframe(st.session_state.df2.head(5), width='stretch')

if st.session_state.df1 is not None and st.session_state.df2 is not None:
    st.info(f"Comparison mode: **{st.session_state.name1}** vs **{st.session_state.name2}**")
elif st.session_state.df1 is not None:
    st.info(f"Single file mode: **{st.session_state.name1}**")


st.divider()

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        if msg["role"] == "assistant":
            st.write(msg["answer"])
            if msg.get("findings"):
                st.subheader("Key Findings")
                for f in msg["findings"]:
                    st.markdown(f"- {f}")
            if msg.get("chart") and os.path.exists(msg["chart"]):
                st.image(msg["chart"], width=800)
            # Re-render validation download button from history
            if msg.get("validation_report_path") and os.path.exists(msg["validation_report_path"]):
                with open(msg["validation_report_path"], "rb") as vf:
                    b64 = base64.b64encode(vf.read()).decode()
                st.markdown(f'''
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                   download="validation_report.xlsx" class="export-dl-btn export-dl-excel">
                   📊 &nbsp; Download Validation Report
                </a>''', unsafe_allow_html=True)
        else:
            st.write(msg["content"])

question = st.chat_input("Ask a question about your data...", key="main_chat")

# ── Quick action chips — shown only when data is loaded and no messages yet ──
if st.session_state.df1 is not None and not question and len(st.session_state.messages) == 0:
    _spacer1, _chip1, _chip2, _spacer2 = st.columns([1.5, 1, 1, 1.5])
    with _chip1:
        if st.button("🔍 Validate this data", key="chip_validate", use_container_width=True):
            st.session_state._chip_query = "Validate this data"
            st.rerun()
    with _chip2:
        if st.button("📊 Analyze this data", key="chip_analyze", use_container_width=True):
            st.session_state._chip_query = "Give me a full analysis of this dataset"
            st.rerun()

# Pick up chip click from previous rerun
if "_chip_query" in st.session_state and st.session_state._chip_query:
    question = st.session_state._chip_query
    st.session_state._chip_query = None

# ── Validation trigger keywords ──
import re as _re
_VALIDATE_PATTERN = _re.compile(
    r"\b(validat\w*|"
    r"check\s+(for\s+)?(error|issue|qualit|duplicat|accura)\w*|"
    r"run\s+validation|data\s+quality|error\s+(rate|check|detect)\w*|"
    r"find\s+(error|issue|duplicat|invalid)\w*|"
    r"duplicate\s+(check|detect)\w*|"
    r"clean\s+check|audit|inspect\s+data|verify\s+(data|this))\b",
    _re.IGNORECASE
)

if question:
    if st.session_state.df1 is None:
        st.warning("Please upload at least one file first.")
    elif _VALIDATE_PATTERN.search(question):
        # ── VALIDATION MODE — no AI call, pure local validation ──
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.write(question)
        with st.chat_message("assistant"):
            with st.spinner("Running validation checks across 13 functions..."):
                vr = run_full_validation(st.session_state.df1)
                st.session_state.validation_results = vr

                # Generate summary text
                summary = validation_summary_text(vr)

                # Export to Excel
                report_path = f"validation_report_{len(st.session_state.messages)}.xlsx"
                export_validation_report(st.session_state.df1, vr, output_path=report_path)

            # Display summary in chat
            st.write("**Data Validation Complete**")
            st.text(summary)

            # Show score cards inline
            overall = vr.get("overall_status", {})
            accuracy = vr.get("data_accuracy", {})
            err_rate_data = vr.get("error_rate", {})
            workflow = vr.get("workflow_status", {})

            score = overall.get("score", 0)
            status_label = overall.get("status", "N/A")
            if status_label == "CLEAN":
                score_color = "#16a34a"
            elif status_label == "WARNINGS":
                score_color = "#d97706"
            else:
                score_color = "#dc2626"

            st.markdown(f"""
            <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:16px 0;">
                <div style="background:{score_color}15; border:1px solid {score_color}40; border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:700; color:{score_color};">{score}%</div>
                    <div style="font-size:12px; color:{score_color}; font-weight:600;">{status_label}</div>
                    <div style="font-size:11px; opacity:0.7; margin-top:4px;">Quality Score</div>
                </div>
                <div style="background:#2563eb15; border:1px solid #2563eb40; border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:700; color:#2563eb;">{accuracy.get('overall_accuracy', 'N/A')}%</div>
                    <div style="font-size:12px; color:#2563eb; font-weight:600;">ACCURACY</div>
                    <div style="font-size:11px; opacity:0.7; margin-top:4px;">Data Accuracy</div>
                </div>
                <div style="background:#dc262615; border:1px solid #dc262640; border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:700; color:#dc2626;">{err_rate_data.get('overall_error_rate', 'N/A')}%</div>
                    <div style="font-size:12px; color:#dc2626; font-weight:600;">ERROR RATE</div>
                    <div style="font-size:11px; opacity:0.7; margin-top:4px;">Overall Errors</div>
                </div>
                <div style="background:#7c3aed15; border:1px solid #7c3aed40; border-radius:12px; padding:16px; text-align:center;">
                    <div style="font-size:28px; font-weight:700; color:#7c3aed;">{workflow.get('completion_rate', 'N/A') if workflow.get('column') else 'N/A'}%</div>
                    <div style="font-size:12px; color:#7c3aed; font-weight:600;">WORKFLOW</div>
                    <div style="font-size:11px; opacity:0.7; margin-top:4px;">Completion Rate</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Download button
            with open(report_path, "rb") as vf:
                b64 = base64.b64encode(vf.read()).decode()
            st.markdown(f'''
            <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
               download="validation_report.xlsx" class="export-dl-btn export-dl-excel">
               📊 &nbsp; Download Validation Report (Excel)
            </a>''', unsafe_allow_html=True)

            st.session_state.messages.append({
                "role": "assistant",
                "answer": "Data Validation Complete\n\n" + summary,
                "findings": vr.get("error_notes", []),
                "chart": None,
                "validation_report_path": report_path,
            })
    else:
        st.session_state.messages.append({"role": "user", "content": question})
        with st.chat_message("user"):
            st.write(question)
        with st.chat_message("assistant"):
            with st.spinner("Analyzing... (this may take up to 30s if servers are busy)"):
                comparing = st.session_state.df2 is not None
                if comparing:
                    analysis = analyze_two_datasets(
                        st.session_state.df1, st.session_state.name1,
                        st.session_state.df2, st.session_state.name2,
                        question, st.session_state.messages
                    )
                    run_analysis_code(
                        analysis.get("python_code", ""),
                        df1=st.session_state.df1,
                        df2=st.session_state.df2
                    )
                else:
                    analysis = analyze_data(
                        st.session_state.df1, question, st.session_state.messages
                    )
                    run_analysis_code(
                        analysis.get("python_code", ""),
                        df=st.session_state.df1
                    )
                st.session_state.last_analysis = analysis

            st.write(analysis.get("answer", ""))
            findings = analysis.get("key_findings", [])
            if findings:
                st.subheader("Key Findings")
                for f in findings:
                    st.markdown(f"- {f}")

            chart_path = f"chart_{len(st.session_state.messages)}.png"
            if os.path.exists("chart.png"):
                shutil.copy("chart.png", chart_path)
                st.image(chart_path, width=800)
            else:
                chart_path = None

            st.session_state.messages.append({
                "role": "assistant",
                "answer": analysis.get("answer", ""),
                "findings": findings,
                "chart": chart_path
            })

        if st.session_state.last_analysis:
            st.markdown("""
            <div class="export-section">
                <p class="export-title">📤 Export Report</p>
            </div>
            """, unsafe_allow_html=True)

            c1, c2, c3 = st.columns(3)
            with c1:
                ep = export_to_excel(
                    st.session_state.df1, st.session_state.last_analysis, "chart.png",
                    df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(ep, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                st.markdown(f'''
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
                   download="report.xlsx" class="export-dl-btn export-dl-excel">
                   📊 &nbsp; Excel Report
                </a>''', unsafe_allow_html=True)
            with c2:
                wp = export_to_word(
                    st.session_state.df1, st.session_state.last_analysis, "chart.png",
                    df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(wp, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                st.markdown(f'''
                <a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}"
                   download="report.docx" class="export-dl-btn export-dl-word">
                   📄 &nbsp; Word Report
                </a>''', unsafe_allow_html=True)
            with c3:
                pp = export_to_pptx(
                    st.session_state.last_analysis, "chart.png",
                    df=st.session_state.df1, df2=st.session_state.df2,
                    name1=st.session_state.name1, name2=st.session_state.name2
                )
                with open(pp, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                st.markdown(f'''
                <a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}"
                   download="report.pptx" class="export-dl-btn export-dl-pptx">
                   📋 &nbsp; PowerPoint Report
                </a>''', unsafe_allow_html=True)

if st.session_state.messages:
    if st.button("Clear Chat", key="clear_chat"):
        st.session_state.messages = []
        st.session_state.last_analysis = None
        st.rerun()
