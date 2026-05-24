import streamlit as st
from ferramentas.utilidades import ir_para
from modulos import dcs

st.set_page_config(page_title="Sistema Comportamento", layout="wide")

# --- 1. INICIALIZAÇÃO DE VARIÁVEIS GLOBAIS (MEMÓRIA) ---
if "pptx_gerado" not in st.session_state: st.session_state.pptx_gerado = None
if "nome_arquivo" not in st.session_state: st.session_state.nome_arquivo = ""
if 'tentou_gerar' not in st.session_state: st.session_state.tentou_gerar = False

if 'current_page' not in st.session_state: 
    st.session_state.current_page = "🏠 Início"

if 'unidades_dcs' not in st.session_state: 
    st.session_state.unidades_dcs = [{"id": 0, "nome": "Unidade Matriz", "pop_total": 0, "lideres": 0}]

if 'fases_dcs' not in st.session_state:
    st.session_state.fases_dcs = [
        {"id": 0, "nome": "Planejamento / Reunião de Abertura", "horas": 4},
        {"id": 1, "nome": "Workshop - Gestão de Cultura de Segurança", "horas": 2},
        {"id": 2, "nome": "Elaboração do Relatório", "horas": 56},
        {"id": 3, "nome": "Apresentação dos Resultados", "horas": 4},
        {"id": 4, "nome": "Plano de Transformação Cultural (PTC)", "horas": 12},
        {"id": 5, "nome": "Suporte e Acompanhamento", "horas": 64}
    ]

if 'memoria_geral' not in st.session_state:
    st.session_state.memoria_geral = {
        "cliente": "", "unidade": "", "num_prop": "", "escopo": "", "prazo": "", 
        "formato": "Híbrido", "justificativa": "", "objetivo": "", "idas": 0
    }

if 'valores_finais' not in st.session_state:
    st.session_state.valores_finais = {"op1": 0.0, "op2": 0.0}


# --- 2. MENU LATERAL ---
st.sidebar.title("🧭 Navegação Integrada")
menu_options = ["🏠 Início", "💰 1. Precificação", "📝 2. Proposta Técnica", "📈 3. Proposta Comercial"]

# A MÁGICA ACONTECE AQUI: O parâmetro 'key' liga o menu direto à memória global!
st.sidebar.radio("Etapa do Projeto:", menu_options, key="current_page")


# --- 3. ROTEADOR DE PÁGINAS ---
if st.session_state.current_page == "🏠 Início":
    st.title("Bem-vindo ao Sistema Integrado - Comportamento")
    st.write("Siga o fluxo integrado: calcule os custos, gere a proposta técnica e, em seguida, a comercial. Os dados acompanham-no!")
    st.button("🚀 Iniciar Novo Projeto (Ir para Precificação)", on_click=ir_para, args=("💰 1. Precificação",))

elif st.session_state.current_page == "💰 1. Precificação":
    dcs.render_precificacao()

elif st.session_state.current_page == "📝 2. Proposta Técnica":
    dcs.render_tecnica()

elif st.session_state.current_page == "📈 3. Proposta Comercial":
    dcs.render_comercial()
