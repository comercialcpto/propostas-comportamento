import streamlit as st
from ferramentas.utilidades import ir_para
from ferramentas import estilo
from modulos import dcs, pontual, handover

st.set_page_config(page_title="Sistema Comportamento", layout="wide", initial_sidebar_state="expanded")
estilo.aplicar_estilo()

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

# =====================================================================
# CARTÕES DA HOME (vitrine de serviços)
# ---------------------------------------------------------------------
# Cada cartão é puramente visual; quando tem 'servico' apontando para uma
# entrada DISPONÍVEL do SERVICOS acima, ele vira "Construir" e clicar leva
# direto à Precificação daquele serviço. Caso contrário, fica "Em breve".
#
# Lever opcional por cartão:
#   "forcar": "em_breve"  -> trava como "Em breve" mesmo com o módulo pronto
#   "forcar": "construir" -> força "Construir" (use SÓ se o módulo existir;
#                            senão o botão cai no aviso "em desenvolvimento")
#
# 'icone' referencia ferramentas/estilo.ICONES_SERVICO.
# =====================================================================
CARDS_HOME = [
    {"id": "dcs", "icone": "dcs",
     "titulo": "Diagnóstico de Cultura de Segurança",
     "subtitulo": "DCS — Hearts & Minds, Plano de Transformação Cultural",
     "servico": "Diagnóstico (DCS/Clima/DCMA)"},

    {"id": "mpl", "icone": "mpl",
     "titulo": "Mapeamento de Perfil de Liderança",
     "subtitulo": "MPL — Desenvolvimento individual e coletivo da liderança",
     "servico": "Mapeamento de Liderança (MPL)"},  # módulo ainda não existe -> Em breve

    {"id": "clima", "icone": "clima",
     "titulo": "Diagnóstico de Clima de Segurança",
     "subtitulo": "Pesquisa de Valência — PREVACC",
     "servico": None},

    {"id": "workshop", "icone": "workshop",
     "titulo": "Workshop & Palestra",
     "subtitulo": "Formações pontuais e eventos",
     "servico": "Proposta Pontual (Palestras/Workshops)",
     "forcar": "em_breve"},  # módulo Pontual JÁ existe; remova esta linha p/ liberar

    {"id": "ativadores", "icone": "ativadores",
     "titulo": "Gestão de Ativadores",
     "subtitulo": "Mapeamento comportamental",
     "servico": None},

    {"id": "outros", "icone": "outros",
     "titulo": "Outros Projetos",
     "subtitulo": "Mentorias e personalizados",
     "servico": None},
]

# Rótulos limpos para a navegação (o VALOR continua sendo a string original,
# para não quebrar as referências a current_page espalhadas nos módulos).
PAGINAS = {
    "🏠 Início": "Início",
    "💰 1. Precificação": "1 · Precificação",
    "📝 2. Proposta Técnica": "2 · Proposta Técnica",
    "📈 3. Proposta Comercial": "3 · Proposta Comercial",
    "🤝 4. Handover (Operações)": "4 · Handover",
}


def servico_atual():
    return st.session_state.get("servico_selecionado", list(SERVICOS.keys())[0])


def modulo_atual():
    info = SERVICOS.get(servico_atual())
    if info and info["disponivel"]:
        return info["modulo"]
    return None


def card_disponivel(card):
    """Cartão é 'Construir' quando aponta para um serviço disponível (ou é forçado)."""
    forcado = card.get("forcar")
    if forcado in ("construir", "em_breve"):
        return forcado == "construir"
    serv = card.get("servico")
    return bool(serv) and SERVICOS.get(serv, {}).get("disponivel", False)


def abrir_servico(servico_key):
    """Callback do cartão: seleciona o serviço e vai para a Precificação."""
    st.session_state.servico_selecionado = servico_key
    ir_para("💰 1. Precificação")


# --- 1. INICIALIZAÇÃO DE VARIÁVEIS GLOBAIS (MEMÓRIA) ---
if "pptx_gerado" not in st.session_state: st.session_state.pptx_gerado = None
if "nome_arquivo" not in st.session_state: st.session_state.nome_arquivo = ""
if 'tentou_gerar' not in st.session_state: st.session_state.tentou_gerar = False

if 'current_page' not in st.session_state:
    st.session_state.current_page = "🏠 Início"

# --- Memória DCS ---
if 'unidades_dcs' not in st.session_state:
    st.session_state.unidades_dcs = [
        {"id": 0, "nome": "Unidade Matriz", "pop_total": 0, "lideres": 0,
         "oac": 0, "visitas": 0, "aprofundamento": 0}
    ]

if 'fases_dcs' not in st.session_state:
    st.session_state.fases_dcs = [
        {"id": 0, "nome": "Planejamento / Reunião de Abertura", "horas": 4, "presencial": False},
        {"id": 1, "nome": "Workshop - Gestão de Cultura de Segurança", "horas": 2, "presencial": False},
        {"id": 2, "nome": "Elaboração do Relatório", "horas": 56, "presencial": False},
        {"id": 3, "nome": "Apresentação dos Resultados", "horas": 4, "presencial": False},
        {"id": 4, "nome": "Plano de Transformação Cultural (PTC)", "horas": 12, "presencial": False},
        {"id": 5, "nome": "Suporte e Acompanhamento", "horas": 64, "presencial": False}
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
estilo.marca_sidebar()
st.sidebar.radio(
    "Navegação",
    list(PAGINAS.keys()),
    key="current_page",
    format_func=lambda x: PAGINAS[x],
    label_visibility="collapsed",
)


# --- 3. ROTEADOR DE PÁGINAS ---
if st.session_state.current_page == "🏠 Início":
    st.markdown(
        '<div class="cpto-home-head">'
        '<div class="cpto-eyebrow">Grupo Comportamento</div>'
        '<h1>Gerador de Propostas</h1>'
        '<p>Selecione o tipo de proposta para começar</p>'
        '</div>',
        unsafe_allow_html=True,
    )

    for inicio in range(0, len(CARDS_HOME), 3):
        colunas = st.columns(3, gap="medium")
        for coluna, card in zip(colunas, CARDS_HOME[inicio:inicio + 3]):
            with coluna:
                disponivel = card_disponivel(card)
                estilo.card_servico(card, disponivel)
                if disponivel:
                    st.button(
                        "Construir",
                        key=f"abrir_{card['id']}",
                        type="primary",
                        use_container_width=True,
                        on_click=abrir_servico,
                        args=(card["servico"],),
                    )

elif st.session_state.current_page == "💰 1. Precificação":
    estilo.cabecalho("Precificação", "Escolha o serviço e construa o investimento com o racional aberto.", etapa="Etapa 1")

    servico = st.selectbox(
        "Serviço",
        list(SERVICOS.keys()),
        key="servico_selecionado",
    )
    st.markdown("---")

    info = SERVICOS[servico]
    if info["disponivel"]:
        info["modulo"].render_precificacao()
    else:
        st.info("Módulo em desenvolvimento. Em breve disponível por aqui.")

elif st.session_state.current_page == "📝 2. Proposta Técnica":
    mod = modulo_atual()
    if mod:
        mod.render_tecnica()
    else:
        st.info("Este serviço ainda não tem Proposta Técnica implementada.")

elif st.session_state.current_page == "📈 3. Proposta Comercial":
    mod = modulo_atual()
    if mod:
        mod.render_comercial()
    else:
        st.info("Este serviço ainda não tem Proposta Comercial implementada.")

elif st.session_state.current_page == "🤝 4. Handover (Operações)":
    handover.render_handover()
