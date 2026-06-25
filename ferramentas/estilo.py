"""
Camada de identidade visual do Sistema Comportamento — modo ESCURO.

TODA a customização visual fica centralizada aqui (e nas cores do
.streamlit/config.toml).

Paleta:
  navy exato do PPT (fundo do slide #0C2539)  = estrutura
  teal exato da marca (#03857B)               = highlight/hover

Uso no app.py:
    from ferramentas import estilo
    estilo.aplicar_estilo()                 # logo após set_page_config
    estilo.cabecalho("Precificação", "...", etapa="Etapa 1")
    estilo.marca_sidebar()
    estilo.card_servico(card, disponivel)   # cartões da home
"""
import os
import streamlit as st

# Logo opcional. Coloque o PNG (branco, fundo transparente) em assets/.
# Se o arquivo não existir, cai no texto "Comportamento." automaticamente.
LOGO_PATH = "assets/logo_comportamento.png"

_CSS = """
<style>
:root{
  --canvas:#0C2539;        /* fundo: navy exato do slide PPT */
  --surface:#112E46;       /* cartões, inputs */
  --surface-2:#173C57;     /* hover/elevado */
  --sidebar:#0E2A40;
  --line:#244B6B;
  --ink:#EAF1F8;           /* texto principal */
  --slate:#8DA4B9;         /* texto secundário */
  --accent:#03857B;        /* highlight: teal exato da marca */
  --accent-bright:#16B6A6; /* hover/foco */
  --accent-wash:rgba(16,179,164,.16);
  --accent-glow:rgba(22,182,166,.34);
  --shadow:0 1px 2px rgba(0,0,0,.30), 0 10px 30px rgba(0,0,0,.38);
  --ease:cubic-bezier(.2,.7,.2,1.2);
}

/* ---- Canvas e chrome ----
   IMPORTANTE: NÃO esconder o stToolbar inteiro — o botão de reexpandir a
   sidebar (stExpandSidebarButton) é FILHO dele. Escondemos só o ruído. */
.stApp{ background:var(--canvas); }
[data-testid="stDecoration"]{ display:none !important; }
#MainMenu, [data-testid="stMainMenuButton"]{ display:none !important; }
footer{ display:none !important; }
header[data-testid="stHeader"]{ background:transparent; }
[data-testid="stToolbar"]{ background:transparent; }
[data-testid="stMainBlockContainer"]{ max-width:1240px; padding-top:2.2rem; padding-bottom:4rem; }

/* Botões de recolher/expandir a sidebar SEMPRE visíveis e clicáveis */
[data-testid="stExpandSidebarButton"], [data-testid="stExpandSidebarButton"] button,
[data-testid="stSidebarCollapseButton"], [data-testid="stSidebarCollapseButton"] button{
  display:inline-flex !important; visibility:visible !important; opacity:1 !important;
  pointer-events:auto !important; z-index:1000000 !important; color:var(--ink) !important;
}

/* ---- Tipografia ---- */
h1,h2,h3{ color:var(--ink); letter-spacing:-.012em; }
h1{ font-weight:700; font-size:1.9rem; }
h2{ font-weight:600; font-size:1.35rem; }
h3{ font-weight:600; font-size:1.12rem; }
p, span, label, li, .stMarkdown{ color:var(--ink); }
hr{ border:none; border-top:1px solid var(--line); margin:1.4rem 0; }

/* ---- Cabeçalho de página ---- */
.cpto-head{ margin-bottom:1.5rem; }
.cpto-eyebrow{
  text-transform:uppercase; letter-spacing:.14em; font-size:.7rem; font-weight:600;
  color:var(--accent-bright); font-family:'Space Grotesk',sans-serif; margin-bottom:.4rem;
}
.cpto-title{ margin:0; font-size:1.9rem; font-weight:700; color:var(--ink); }
.cpto-sub{ color:var(--slate); margin:.45rem 0 0; font-size:.97rem; max-width:62ch; }

/* ---- Home: cabeçalho central + grade de serviços ---- */
.cpto-home-head{ text-align:center; margin:.4rem 0 2.2rem; }
.cpto-home-head .cpto-eyebrow{ margin-bottom:.55rem; }
.cpto-home-head h1{
  margin:0; font-size:2.4rem; font-weight:700; letter-spacing:-.02em; color:var(--ink);
}
.cpto-home-head p{ color:var(--slate); margin:.7rem auto 0; max-width:52ch; font-size:1rem; }

.cpto-card{
  display:flex; flex-direction:column;
  background:var(--surface); border:1px solid var(--line);
  border-radius:1rem; padding:1.25rem 1.2rem 1.2rem; min-height:186px;
  transition:border-color .16s ease, transform .16s var(--ease), box-shadow .16s ease;
}
.cpto-card:not(.cpto-card--soon):hover{
  border-color:var(--accent); transform:translateY(-2px);
  box-shadow:0 12px 28px rgba(0,0,0,.35);
}
.cpto-card--soon{ opacity:.55; }
.cpto-card-ic{
  width:46px; height:46px; display:grid; place-items:center; border-radius:.78rem;
  background:var(--accent-wash); color:var(--accent-bright); margin-bottom:.95rem;
}
.cpto-card--soon .cpto-card-ic{ background:rgba(141,164,185,.10); color:var(--slate); }
.cpto-card-ic svg{ width:23px; height:23px; }
.cpto-card h3{ font-size:1.02rem; font-weight:600; margin:0 0 .35rem; line-height:1.25; color:var(--ink); }
.cpto-card p{ color:var(--slate); font-size:.86rem; line-height:1.45; margin:0; }
.cpto-pill{ margin-top:auto; padding-top:1rem; }
.cpto-pill span{
  font-family:'Space Grotesk',sans-serif; font-weight:600; font-size:.66rem;
  letter-spacing:.14em; text-transform:uppercase; padding:.34rem .72rem;
  border-radius:999px; display:inline-block;
  color:var(--slate); background:rgba(141,164,185,.08); border:1px solid rgba(141,164,185,.22);
}

/* ---- Sidebar ---- */
[data-testid="stSidebar"]{ border-right:1px solid var(--line); }
[data-testid="stSidebar"] > div{ padding-top:1.4rem; }
[data-testid="stSidebar"] [data-testid="stImage"]{ margin:.1rem 0 .5rem; }
.cpto-brand{ padding:0 .35rem 1.1rem; }
.cpto-brand .mark{
  font-family:'Space Grotesk',sans-serif; font-weight:700; font-size:1.05rem;
  color:var(--ink); letter-spacing:-.01em;
}
.cpto-brand .dot{ color:var(--accent-bright); }
.cpto-brand .tag{ color:var(--slate); font-size:.78rem; margin-top:.15rem; }
.cpto-navlabel{
  text-transform:uppercase; letter-spacing:.13em; font-size:.66rem; font-weight:600;
  color:var(--slate); padding:0 .35rem .5rem;
}

/* Trilha de etapas: o st.radio vira navegação de app, com hover que desliza */
[data-testid="stSidebar"] div[role="radiogroup"]{ gap:.18rem; }
[data-testid="stSidebar"] div[role="radiogroup"] > label{
  display:flex; align-items:center; padding:.55rem .7rem; margin:0;
  border-radius:.6rem; cursor:pointer; color:var(--slate);
  font-weight:500; font-size:.93rem; line-height:1.2;
  transition:background .14s ease, color .14s ease, box-shadow .14s ease, transform .14s var(--ease);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label:hover{
  background:var(--surface-2); color:var(--ink); transform:translateX(3px);
}
[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child{ display:none; }
[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked){
  background:var(--accent-wash); color:#fff; font-weight:600;
  box-shadow:inset 3px 0 0 var(--accent-bright);
}

/* ---- Botões (saltam no hover) ---- */
.stButton > button, .stDownloadButton > button{
  border-radius:.6rem; font-weight:600; font-size:.92rem; padding:.5rem 1.05rem;
  border:1px solid var(--line); background:var(--surface); color:var(--ink);
  box-shadow:0 1px 1px rgba(0,0,0,.25);
  transition:transform .18s var(--ease), box-shadow .18s ease, background .15s ease, border-color .15s ease;
}
.stButton > button:hover, .stDownloadButton > button:hover{
  transform:translateY(-2px); border-color:var(--accent);
  box-shadow:0 8px 20px rgba(0,0,0,.4), 0 0 0 1px var(--accent-glow);
}
.stButton > button:active, .stDownloadButton > button:active{ transform:translateY(0); }
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"]{
  background:var(--accent); border-color:var(--accent); color:#fff;
}
.stButton > button[kind="primary"]:hover{
  background:var(--accent-bright); border-color:var(--accent-bright);
  box-shadow:0 10px 24px var(--accent-glow);
}
.stButton > button:focus-visible, .stDownloadButton > button:focus-visible{
  outline:2px solid var(--accent-bright); outline-offset:2px;
}

/* ---- Inputs (highlight no hover, glow no foco) ---- */
.stTextInput input, .stNumberInput input, .stTextArea textarea{
  border-radius:.55rem !important; border:1px solid var(--line);
  background:var(--surface); color:var(--ink);
  transition:border-color .15s ease, box-shadow .15s ease, background .15s ease;
}
.stTextInput input:hover, .stNumberInput input:hover, .stTextArea textarea:hover{
  border-color:var(--accent); background:var(--surface-2);
}
.stTextInput input:focus, .stNumberInput input:focus, .stTextArea textarea:focus{
  border-color:var(--accent); background:var(--surface-2);
  box-shadow:0 0 0 3px var(--accent-wash);
}
.stSelectbox div[data-baseweb="select"] > div, .stMultiSelect div[data-baseweb="select"] > div{
  border-radius:.55rem; border-color:var(--line); background:var(--surface);
  transition:border-color .15s ease, background .15s ease, box-shadow .15s ease;
}
.stSelectbox div[data-baseweb="select"] > div:hover,
.stMultiSelect div[data-baseweb="select"] > div:hover{
  border-color:var(--accent); background:var(--surface-2);
}

/* ---- Métricas como cartões ---- */
[data-testid="stMetric"]{
  background:var(--surface); border:1px solid var(--line); border-radius:.85rem;
  padding:1rem 1.15rem; box-shadow:var(--shadow);
  transition:border-color .15s ease, transform .15s var(--ease), box-shadow .15s ease;
}
[data-testid="stMetric"]:hover{ border-color:var(--accent); transform:translateY(-2px); }
[data-testid="stMetricValue"]{ font-family:'Space Grotesk',sans-serif; color:var(--ink); }
[data-testid="stMetricLabel"] p{ color:var(--slate); font-weight:500; }

/* ---- Expanders como cartões ---- */
[data-testid="stExpander"]{
  border:1px solid var(--line); border-radius:.85rem; background:var(--surface);
  box-shadow:var(--shadow); overflow:hidden;
  transition:border-color .15s ease;
}
[data-testid="stExpander"]:hover{ border-color:var(--accent); }
[data-testid="stExpander"] summary{ padding:.85rem 1.05rem; font-weight:600; color:var(--ink); }
[data-testid="stExpander"] summary:hover{ color:var(--accent-bright); }

/* ---- Alerts / tabelas ---- */
[data-testid="stAlert"]{ border-radius:.7rem; }
[data-testid="stTable"] table{ border-radius:.6rem; overflow:hidden; }
[data-testid="stTable"] thead th{ background:var(--surface-2); color:var(--ink); font-weight:600; }
[data-testid="stTable"] tbody td{ color:var(--ink); }

/* ---- Acessibilidade ---- */
@media (prefers-reduced-motion: reduce){ *{ transition:none !important; } }
</style>
"""

# SVGs dos cartões da home (line icons, herdam a cor do contêiner .cpto-card-ic).
ICONES_SERVICO = {
    "dcs": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/>'
        '<path d="M11 8v6M8 11h6"/></svg>'
    ),
    "mpl": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<circle cx="9" cy="9" r="3.2"/><path d="M3 20v-0.5a6 6 0 0 1 12 0V20"/>'
        '<g transform="translate(14.5 1.2) scale(0.36)" stroke="none" fill="currentColor">'
        '<path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14l-5-4.87 '
        '6.91-1.01L12 2z"/></g></svg>'
    ),
    "clima": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M14 14.76V5a2 2 0 1 0-4 0v9.76a4 4 0 1 0 4 0z"/></svg>'
    ),
    "workshop": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<rect x="9" y="2" width="6" height="11" rx="3"/>'
        '<path d="M5 10v2a7 7 0 0 0 14 0v-2"/><path d="M12 19v3"/><path d="M8 22h8"/></svg>'
    ),
    "ativadores": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M2 5a2 2 0 0 1 2-2h6v16H4a2 2 0 0 0-2 2V5z"/>'
        '<path d="M22 5a2 2 0 0 0-2-2h-6v16h6a2 2 0 0 1 2 2V5z"/></svg>'
    ),
    "outros": (
        '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" '
        'stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>'
        '<path d="M14 2v6h6M9 13h6M9 17h6"/></svg>'
    ),
}


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
    """Logo da marca na barra lateral (com fallback para texto) + rótulo da nav."""
    if os.path.exists(LOGO_PATH):
        st.sidebar.image(LOGO_PATH, width=185)
    else:
        st.sidebar.markdown(
            '<div class="cpto-brand"><div class="mark">Comportamento<span class="dot">.</span></div></div>',
            unsafe_allow_html=True,
        )
    st.sidebar.markdown(
        '<div class="cpto-brand" style="padding:.35rem .35rem 1.05rem;">'
        '<div class="tag">Precificação & Propostas</div></div>'
        f'<div class="cpto-navlabel">{navlabel}</div>',
        unsafe_allow_html=True,
    )


def card_servico(card, disponivel):
    """Renderiza o corpo de um cartão de serviço da home.

    Cartões 'Construir' (disponivel=True) recebem o botão de ação no app.py,
    logo abaixo deste corpo. Cartões 'Em breve' já trazem o selo embutido.
    """
    icone = ICONES_SERVICO.get(card.get("icone", ""), "")
    modificador = "" if disponivel else " cpto-card--soon"
    selo = "" if disponivel else '<div class="cpto-pill"><span>Em breve</span></div>'
    st.markdown(
        f'<div class="cpto-card{modificador}">'
        f'<div class="cpto-card-ic">{icone}</div>'
        f'<h3>{card.get("titulo", "")}</h3>'
        f'<p>{card.get("subtitulo", "")}</p>'
        f'{selo}'
        f'</div>',
        unsafe_allow_html=True,
    )
