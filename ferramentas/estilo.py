"""
Camada de identidade visual do Sistema Comportamento.

TODA a customização visual fica centralizada aqui (e nas cores do
.streamlit/config.toml). Se um dia o Streamlit mudar algum nome de classe
interna e algo "descolar", o conserto é só neste arquivo — nunca espalhado
pelos módulos.

Uso no app.py:
    from ferramentas import estilo
    estilo.aplicar_estilo()                 # logo após set_page_config
    estilo.cabecalho("Precificação", "...", etapa="Etapa 1")
"""
import streamlit as st

_CSS = """
<style>
:root{
  --ink:#0F1B2D; --slate:#5A6B7B; --line:#E3E8EE; --surface:#FFFFFF;
  --canvas:#F4F6F8; --accent:#009974; --accent-dark:#007A5C;
  --accent-wash:rgba(0,153,116,.10);
  --shadow:0 1px 2px rgba(15,27,45,.04), 0 8px 24px rgba(15,27,45,.06);
}

/* ---- Canvas e chrome ---- */
.stApp{ background:var(--canvas); }
[data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer{ display:none !important; }
header[data-testid="stHeader"]{ background:transparent; }
[data-testid="stMainBlockContainer"]{ max-width:1240px; padding-top:2.2rem; padding-bottom:4rem; }

/* ---- Tipografia ---- */
h1,h2,h3{ color:var(--ink); letter-spacing:-.012em; }
h1{ font-weight:700; font-size:1.9rem; }
h2{ font-weight:600; font-size:1.35rem; }
h3{ font-weight:600; font-size:1.12rem; }
hr{ border:none; border-top:1px solid var(--line); margin:1.4rem 0; }

/* ---- Cabeçalho de página (helper cabecalho()) ---- */
.cpto-head{ margin-bottom:1.5rem; }
.cpto-eyebrow{
  text-transform:uppercase; letter-spacing:.14em; font-size:.7rem; font-weight:600;
  color:var(--accent); font-family:'Space Grotesk',sans-serif; margin-bottom:.4rem;
}
.cpto-title{ margin:0; font-size:1.9rem; font-weight:700; color:var(--ink); }
.cpto-sub{ color:var(--slate); margin:.45rem 0 0; font-size:.97rem; max-width:62ch; }

/* ---- Sidebar ---- */
[data-testid="stSidebar"]{ border-right:1px solid var(--line); }
[data-testid="stSidebar"] > div{ padding-top:1.4rem; }
.cpto-brand{ padding:0 .35rem 1.1rem; }
.cpto-brand .mark{
  font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:1.05rem;
  color:var(--ink); letter-spacing:-.01em;
}
.cpto-brand .dot{ color:var(--accent); }
.cpto-brand .tag{ color:var(--slate); font-size:.78rem; margin-top:.15rem; }
.cpto-navlabel{
  text-transform:uppercase; letter-spacing:.13em; font-size:.66rem; font-weight:600;
  color:var(--slate); padding:0 .35rem .5rem;
}

/* Trilha de etapas: o st.radio vira navegação de app */
[data-testid="stSidebar"] div[role="radiogroup"]{ gap:.18rem; }
[data-testid="stSidebar"] div[role="radiogroup"] > label{
  display:flex; align-items:center; padding:.55rem .7rem; margin:0;
  border-radius:.6rem; cursor:pointer; color:var(--slate);
  font-weight:500; font-size:.93rem; line-height:1.2;
  transition:background .12s ease, color .12s ease, box-shadow .12s ease;
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{ background:var(--canvas); color:var(--ink); }
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child{ display:none; }  /* esconde o círculo */
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){
  background:var(--accent-wash); color:var(--accent-dark); font-weight:600;
  box-shadow:inset 3px 0 0 var(--accent);
}

/* ---- Botões ---- */
.stButton > button, .stDownloadButton > button{
  border-radius:.6rem; font-weight:600; font-size:.92rem; padding:.5rem 1.05rem;
  border:1px solid var(--line); background:var(--surface); color:var(--ink);
  box-shadow:0 1px 1px rgba(15,27,45,.03);
  transition:transform .1s ease, box-shadow .12s ease, background .12s ease, border-color .12s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover{
  border-color:#C9D2DC; transform:translateY(-1px); box-shadow:var(--shadow);
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background:var(--accent); border-color:var(--accent); color:#fff;
}
.stButton > button[kind="primary"]:hover{ background:var(--accent-dark); border-color:var(--accent-dark); }
.stButton > button:focus-visible, .stDownloadButton > button:focus-visible{
  outline:2px solid var(--accent); outline-offset:2px;
}

/* ---- Inputs ---- */
.stTextInput input, .stNumberInput input, .stTextArea textarea{
  border-radius:.55rem !important; border:1px solid var(--line); background:var(--surface);
}
.stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus{
  border-color:var(--accent); box-shadow:0 0 0 3px var(--accent-wash);
}
.stSelectbox div[data-baseweb="select"] > div, .stMultiSelect div[data-baseweb="select"] > div{
  border-radius:.55rem; border-color:var(--line); background:var(--surface);
}

/* ---- Métricas como cartões ---- */
[data-testid="stMetric"]{
  background:var(--surface); border:1px solid var(--line); border-radius:.85rem;
  padding:1rem 1.15rem; box-shadow:var(--shadow);
}
[data-testid="stMetricValue"]{ font-family:'Space Grotesk',sans-serif; color:var(--ink); }
[data-testid="stMetricLabel"] p{ color:var(--slate); font-weight:500; }

/* ---- Expanders como cartões ---- */
[data-testid="stExpander"]{
  border:1px solid var(--line); border-radius:.85rem; background:var(--surface);
  box-shadow:var(--shadow); overflow:hidden;
}
[data-testid="stExpander"] summary{ padding:.85rem 1.05rem; font-weight:600; color:var(--ink); }
[data-testid="stExpander"] summary:hover{ color:var(--accent-dark); }

/* ---- Alerts / tabelas ---- */
[data-testid="stAlert"]{ border-radius:.7rem; }
[data-testid="stTable"] table{ border-radius:.6rem; overflow:hidden; }
[data-testid="stTable"] thead th{ background:var(--canvas); color:var(--ink); font-weight:600; }

/* ---- Acessibilidade ---- */
@media (prefers-reduced-motion: reduce){ *{ transition:none !important; } }
</style>
"""


def aplicar_estilo():
    """Injeta a identidade visual. Chamar UMA vez, logo após set_page_config."""
    st.markdown(_CSS, unsafe_allow_html=True)


def cabecalho(titulo, subtitulo=None, etapa=None):
    """Cabeçalho de página consistente (eyebrow + título + subtítulo)."""
    eyebrow = f'<div class="cpto-eyebrow">{etapa}</div>' if etapa else ""
    sub = f'<p class="cpto-sub">{subtitulo}</p>' if subtitulo else ""
    st.markdown(
        f'<div class="cpto-head">{eyebrow}<h1 class="cpto-title">{titulo}</h1>{sub}</div>',
        unsafe_allow_html=True,
    )


def marca_sidebar(navlabel="Etapas do projeto"):
    """Cabeçalho de marca da barra lateral + rótulo da navegação."""
    st.sidebar.markdown(
        '<div class="cpto-brand">'
        '<div class="mark">Comportamento<span class="dot">.</span></div>'
        '<div class="tag">Precificação & Propostas</div>'
        '</div>'
        f'<div class="cpto-navlabel">{navlabel}</div>',
        unsafe_allow_html=True,
    )
