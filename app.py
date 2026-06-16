import streamlit as st
from ferramentas.utilidades import ir_para
from modulos import dcs, pontual, handover

st.set_page_config(page_title="Sistema Comportamento", layout="wide")

# =====================================================================
# REGISTRO CENTRAL DE SERVIÇOS
# ---------------------------------------------------------------------
# Para adicionar um módulo novo (MAC, RPS, Pulse...), basta:
#   1. Criar modulos/novo.py com render_precificacao/tecnica/comercial
#   2. Importar acima e adicionar UMA entrada aqui.
# O 'template' já fica registrado para a futura troca dinâmica de PPTX.
# =====================================================================
SERVICOS = {
    "Diagnóstico (DCS/Clima/DCMA)": {
        "modulo": dcs, "template": "template_dcs.pptx", "disponivel": True,
    },
    "Proposta Pontual (Palestras/Workshops)": {
        "modulo": pontual, "template": "template_pontual.pptx", "disponivel": True,
    },
    "Mapeamento de Liderança (MPL)": {
        "modulo": None, "template": "template_mpl.pptx", "disponivel": False,
    },
}


def servico_atual():
    return st.session_state.get("servico_selecionado", list(SERVICOS.keys())[0])


def modulo_atual():
    info = SERVICOS.get(servico_atual())
    if info and info["disponivel"]:
        return info["modulo"]
    return None


# --- 1. INICIALIZAÇÃO DE VARIÁVEIS GLOBAIS (MEMÓRIA) ---
if "pptx_gerado" not in st.session_state: st.session_state.pptx_gerado = None
if "nome_arquivo" not in st.session_state: st.session_state.nome_arquivo = ""
if 'tentou_gerar' not in st.session_state: st.session_state.tentou_gerar = False

if 'current_page' not in st.session_state:
    st.session_state.current_page = "🏠 Início"

# --- Memória DCS ---
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

# --- Memória Pontual ---
if 'unidades_pontual' not in st.session_state:
    st.session_state.unidades_pontual = [{"id": 0, "nome": "Público Geral", "pop_total": 0, "lideres": 0}]

if 'fases_pontual' not in st.session_state:
    st.session_state.fases_pontual = [
        {"id": 0, "nome": "Abertura de Projeto", "horas": 2},
        {"id": 1, "nome": "Desenvolvimento dos materiais", "horas": 8},
        {"id": 2, "nome": "Palestra ou Workshop", "horas": 2},
        {"id": 3, "nome": "Análise Crítica", "horas": 2}
    ]

if 'carrinho_pontual' not in st.session_state:
    st.session_state.carrinho_pontual = []

# --- Memória Geral e Logística ---
if 'memoria_geral' not in st.session_state:
    st.session_state.memoria_geral = {
        "cliente": "", "unidade": "", "num_prop": "", "escopo": "", "prazo": "",
        "formato": "Híbrido", "justificativa": "", "objetivo": "", "idas": 0, "idioma": "Português"
    }

if 'valores_finais' not in st.session_state:
    st.session_state.valores_finais = {"op1": 0.0, "op2": 0.0, "horas_totais": 0, "taxa_hora": 0.0, "qtd_parcelas": 1}

if 'logistica_dados' not in st.session_state:
    st.session_state.logistica_dados = {"tipo": "", "total": 0.0, "idas": 0, "detalhes": {}}

if 'servico_selecionado' not in st.session_state:
    st.session_state.servico_selecionado = list(SERVICOS.keys())[0]

# --- 2. MENU LATERAL ---
st.sidebar.title("🧭 Navegação Integrada")
menu_options = ["🏠 Início", "💰 1. Precificação", "📝 2. Proposta Técnica", "📈 3. Proposta Comercial", "🤝 4. Handover (Operações)"]
st.sidebar.radio("Etapa do Projeto:", menu_options, key="current_page")


# --- 3. ROTEADOR DE PÁGINAS ---
if st.session_state.current_page == "🏠 Início":
    st.title("Bem-vindo ao Sistema Integrado - Comportamento")
    st.write("Siga o fluxo integrado: calcule os custos, gere a proposta técnica e, em seguida, a comercial. Os dados acompanham-no!")
    st.button("🚀 Iniciar Novo Projeto (Ir para Precificação)", on_click=ir_para, args=("💰 1. Precificação",))

elif st.session_state.current_page == "💰 1. Precificação":
    st.title("💰 1. Motor de Precificação")

    servico = st.selectbox(
        "Selecione o Serviço para Precificar:",
        list(SERVICOS.keys()),
        index=list(SERVICOS.keys()).index(servico_atual())
    )
    st.session_state.servico_selecionado = servico
    st.markdown("---")

    info = SERVICOS[servico]
    if info["disponivel"]:
        info["modulo"].render_precificacao()
    else:
        st.info("🚧 Módulo em desenvolvimento. Em breve disponível por aqui.")

elif st.session_state.current_page == "📝 2. Proposta Técnica":
    mod = modulo_atual()
    if mod:
        mod.render_tecnica()
    else:
        st.info("🚧 Este serviço ainda não tem Proposta Técnica implementada.")

elif st.session_state.current_page == "📈 3. Proposta Comercial":
    mod = modulo_atual()
    if mod:
        mod.render_comercial()
    else:
        st.info("🚧 Este serviço ainda não tem Proposta Comercial implementada.")

elif st.session_state.current_page == "🤝 4. Handover (Operações)":
    handover.render_handover()
