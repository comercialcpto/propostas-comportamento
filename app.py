import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
import io
import datetime

# --- CONFIGURAÇÕES E LÓGICA ---
def calcular_ch_pela_populacao(pop_total):
    if pop_total <= 100: return 108
    elif pop_total <= 200: return 128
    elif pop_total <= 500: return 144
    elif pop_total <= 800: return 160
    else: return 216

# --- INTERFACE ---
st.set_page_config(page_title="Emissor CPTO v4.5", layout="wide")
st.title("🚀 Emissor de Propostas Grupo Comportamento - v4.5")

# --- SIDEBAR FINANCEIRA ---
with st.sidebar:
    st.header("💰 Financeiro")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Imposto:", ["Comportamento (20%)", "Escola (11%)"])
    imposto = 0.20 if "20%" in entidade else 0.11

# --- ABAS DE PREENCHIMENTO ---
tab1, tab2, tab3, tab4 = st.tabs(["📝 Identificação", "👥 Público-Alvo", "📅 Cronograma (Gantt)", "📊 Resultados"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        cliente = st.text_input("Nome da Empresa ({{CLIENTE}})")
        unidade = st.text_input("Unidade ({{UNIDADE}})")
        num_prop = st.text_input("Nº Proposta ({{NUM_PROP}})")
    with col2:
        idioma = st.selectbox("Idioma ({{IDIOMA}})", ["Português", "Espanhol", "Inglês"])
        formato = st.selectbox("Formato ({{FORMATO}})", ["Híbrido", "Presencial", "Online"])
        prazo = st.text_input("Prazo ({{PRAZO}})", "8 meses")
    
    justificativa = st.text_area("Justificativa ({{JUSTIFICATIVA}})")
    objetivo = st.text_area("Objetivo ({{OBJETIVO}})")

with tab2:
    st.subheader("Detalhamento de Colaboradores")
    c1, c2, c3 = st.columns(3)
    n_pr = c1.number_input("Nº Executivos (Pres/VP/Dir)", value=0)
    n_exec = c2.number_input("Nº Alta Liderança (Gerentes)", value=0)
    n_coord = c3.number_input("Nº Coordenadores", value=0)
    n_super = c1.number_input("Nº Supervisores", value=0)
    n_lid = n_pr + n_exec + n_coord + n_super
    n_sec = c2.number_input("Nº Equipe Segurança", value=0)
    n_oper = c3.number_input("Nº Colaboradores não líderes", value=0)
    
    st.markdown("---")
    n_col3 = c1.number_input("Nº Colaboradores Terceiros", value=0)
    n_lid3 = c2.number_input("Nº Líderes Terceiros", value=0)
    
    # Cálculos Automáticos
    n_prop = n_pr + n_exec + n_coord + n_super + n_sec + n_oper
    n_p_terc = n_prop + n_col3 + n_lid3
    
    st.write(f"**Total Próprios ({{N_PROP}}):** {n_prop}")
    st.write(f"**Total Geral ({{N_PTERC}}):** {n_p_terc}")

with tab3:
    st.subheader("Gantt 12x12 (Meses de Avanço)")
    st.write("Defina as atividades e marque os meses (X) para pintar de verde no slide.")
    
    # Gerador de matriz 12x12 simplificado para a interface
    cronograma_data = []
    for i in range(5): # Exemplo com 5 linhas iniciais
        ativ = st.text_input(f"Atividade {i+1}", key=f"at{i}")
        meses = st.multiselect(f"Meses de execução - {ativ}", list(range(1, 13)), key=f"m{i}")
        cronograma_data.append({"ativ": ativ, "meses": meses})

with tab4:
    # Lógica de relatórios
    col_r1, col_r2 = st.columns(2)
    qtd_rel = col_r1.number_input("Qtd de Unidades com relatório ({{QTD_REL}})", value=1)
    tem_corp = col_r2.checkbox("Gerar Relatório Corporativo?")
    tot_rel = qtd_rel + (1 if tem_corp else 0)
    tot_plan = st.number_input("Total de PTCs ({{TOT_PLAN}})", value=1)

    # Cálculo Financeiro
    ch_base = calcular_ch_pela_populacao(n_p_terc)
    investimento = (ch_base * valor_hora) / (1 - imposto)

    st.markdown("---")
    st.metric("Investimento Estimado", f"R$ {investimento:,.2f}")
    
    if st.button("🚀 GERAR PPTX FINAL"):
        # Dicionário de Mapeamento Completo
        mapeamento = {
            "{{CLIENTE}}": cliente,
            "{{UNIDADE}}": unidade,
            "{{NUM_PROP}}": num_prop,
            "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
            "{{JUSTIFICATIVA}}": justificativa,
            "{{OBJETIVO}}": objetivo,
            "{{PUBLICO}}": n_p_terc,
            "{{PRAZO}}": prazo,
            "{{FORMATO}}": formato,
            "{{IDIOMA}}": idioma,
            "{{N_PR}}": n_pr,
            "{{N_EXEC}}": n_exec,
            "{{N_COORD}}": n_coord,
            "{{N_SUPER}}": n_super,
            "{{N_LID}}": n_lid,
            "{{N_SEC}}": n_sec,
            "{{N_OPER}}": n_oper,
            "{{N_PROP}}": n_prop,
            "{{N_COL3}}": n_col3,
            "{{N_LID3}}": n_lid3,
            "{{N_PTERC}}": n_p_terc,
            "{{TOT_REL}}": tot_rel,
            "{{QTD_REL}}": qtd_rel,
            "{{TOT_PLAN}}": tot_plan
        }
        st.success("Pronto para injetar no template!")
        # AQUI ENTRARÁ A FUNÇÃO DE SUBSTITUIÇÃO QUANDO O ARQUIVO FOR SUBIDO
