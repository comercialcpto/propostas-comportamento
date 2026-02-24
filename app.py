import streamlit as st
from pptx import Presentation
import io
import datetime

# --- 1. INTELIGÊNCIA DE NEGÓCIO (LOTES 3 E 4) ---
def calcular_esforço(servico, n_p_terc, n_lideres, tipo_rps, tipo_pontual):
    if servico == "Diagnóstico (DCS/Clima/DCMA)":
        if n_p_terc <= 100: return 108
        elif n_p_terc <= 200: return 128
        elif n_p_terc <= 500: return 144
        elif n_p_terc <= 800: return 160
        else: return 216
    elif servico == "Mapeamento de Liderança (MPL)":
        return (n_lideres * 6) + 20
    elif servico == "Riscos Psicossociais (RPS)":
        return 1072 if "Mapeamento" in tipo_rps else 1606
    elif servico == "Pulse":
        return 56
    elif servico == "Pontuais / Palestras":
        tabela_p = {"Palestra Online": 30, "Palestra Presencial": 36, "Imersão Liderança": 40}
        return tabela_p.get(tipo_pontual, 16)
    return 0

# --- 2. CONFIGURAÇÃO DA INTERFACE ---
st.set_page_config(page_title="Emissor CPTO v5.0", layout="wide")
st.title("🚀 Emissor de Propostas Inteligente - Grupo Comportamento")

# Sidebar para configurações globais
with st.sidebar:
    st.header("⚙️ Parâmetros Financeiros")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Imposto:", ["Comportamento (20%)", "Escola (11%)"])
    imposto = 0.20 if "20%" in entidade else 0.11

# --- 3. FORMULÁRIO DE ENTRADA ---
tab1, tab2, tab3 = st.tabs(["📄 Escopo e Identificação", "👥 Público e Relatórios", "📅 Cronograma Gantt"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        servico = st.selectbox("Selecione o Serviço Principal", [
            "Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)", 
            "Riscos Psicossociais (RPS)", "Pulse", "Pontuais / Palestras"
        ])
        cliente = st.text_input("Nome do Cliente ({{CLIENTE}})")
        unidade = st.text_input("Unidade atendida ({{UNIDADE}})")
        num_prop = st.text_input("Nº da Proposta ({{NUM_PROP}})")
        
    with col2:
        formato = st.selectbox("Formato ({{FORMATO}})", ["Híbrido", "Presencial", "Online"])
        idioma = st.selectbox("Idioma ({{IDIOMA}})", ["Português", "Espanhol", "Inglês"])
        prazo = st.text_input("Prazo de Execução ({{PRAZO}})", "8 meses")
        data_atual = datetime.date.today().strftime("%d/%m/%Y") # {{DATA}}

    justificativa = st.text_area("Justificativa Técnica ({{JUSTIFICATIVA}})")
    objetivo = st.text_area("Objetivo do Projeto ({{OBJETIVO}})")

with tab2:
    st.subheader("Detalhamento do Público-Alvo")
    c1, c2, c3 = st.columns(3)
    # Variáveis hierárquicas conforme solicitado
    n_pr = c1.number_input("N° Executivos (Presidentes/VP/Dir)", value=0)
    n_exec = c2.number_input("N° Alta Liderança (Gerentes)", value=0)
    n_coord = c3.number_input("N° Coordenadores", value=0)
    n_super = c1.number_input("N° Supervisores", value=0)
    n_lid_extra = c2.number_input("Outros Líderes", value=0)
    n_sec = c3.number_input("Equipe Segurança", value=0)
    n_oper = c1.number_input("Colab. não líderes", value=0)
    
    st.markdown("---")
    n_col3 = c2.number_input("Colab. Terceiros", value=0)
    n_lid3 = c3.number_input("Líderes Terceiros", value=0)

    # Cálculos de Soma Automática
    n_lid_total = n_pr + n_exec + n_coord + n_super + n_lid_extra
    n_prop = n_lid_total + n_sec + n_oper
    n_p_terc = n_prop + n_col3 + n_lid3

    st.info(f"Total Próprios: {n_prop} | Total Geral: {n_p_terc}")

    st.subheader("Entregáveis e Relatórios")
    cr1, cr2 = st.columns(2)
    qtd_rel = cr1.number_input("Quantidade de unidades com relatório individual", value=1)
    tem_corp = cr2.checkbox("Gerar um relatório Corporativo consolidado?")
    tot_rel = qtd_rel + (1 if tem_corp else 0)
    tot_plan = st.number_input("Total de Planos de Transformação (PTCs)", value=1)

with tab3:
    st.subheader("Cronograma de Avanço (Gantt 12x12)")
    st.write("Marque os meses em que cada atividade será realizada.")
    atividades = []
    for i in range(6):
        col_at, col_mes = st.columns([0.3, 0.7])
        nome_at = col_at.text_input(f"Atividade {i+1}", key=f"at_name_{i}")
        meses_selecionados = col_mes.multiselect(f"Meses (M1-M12)", list(range(1, 13)), key=f"at_mes_{i}")
        atividades.append({"nome": nome_at, "meses": meses_selecionados})

# --- 4. PROCESSAMENTO FINAL ---
# Inputs condicionais para o cálculo de CH
tipo_rps = ""
tipo_pontual = ""
if servico == "Riscos Psicossociais (RPS)":
    tipo_rps = st.radio("Escopo RPS", ["Mapeamento (5 meses)", "Gestão Completa (17 meses)"])
if servico == "Pontuais / Palestras":
    tipo_pontual = st.selectbox("Tipo de Evento", ["Palestra Online", "Palestra Presencial", "Imersão Liderança"])

ch_calculada = calcular_esforço(servico, n_p_terc, n_lid_total, tipo_rps, tipo_pontual)
investimento = (ch_calculada * valor_hora) / (1 - imposto)

st.markdown("---")
if st.button("🔥 CALCULAR E PREPARAR VARIÁVEIS"):
    st.success(f"Cálculo concluído para {servico}")
    res1, res2, res3 = st.columns(3)
    res1.metric("Carga Horária Total", f"{ch_calculada}h")
    res2.metric("Investimento Global", f"R$ {investimento:,.2f}")
    res3.metric("Público Total", f"{n_p_terc} pessoas")

    # MAPEAMENTO FINAL DE TAGS PARA O PPTX
    mapeamento_tags = {
        "{{CLIENTE}}": cliente, "{{UNIDADE}}": unidade, "{{NUM_PROP}}": num_prop, "{{DATA}}": data_atual,
        "{{JUSTIFICATIVA}}": justificativa, "{{OBJETIVO}}": objetivo, "{{PUBLICO}}": n_p_terc,
        "{{PRAZO}}": prazo, "{{FORMATO}}": formato, "{{IDIOMA}}": idioma,
        "{{N_PR}}": n_pr, "{{N_EXEC}}": n_exec, "{{N_COORD}}": n_coord, "{{N_SUPER}}": n_super,
        "{{N_LID}}": n_lid_total, "{{N_SEC}}": n_sec, "{{N_OPER}}": n_oper, "{{N_PROP}}": n_prop,
        "{{N_COL3}}": n_col3, "{{N_LID3}}": n_lid3, "{{N_PTERC}}": n_p_terc,
        "{{TOT_REL}}": tot_rel, "{{QTD_REL}}": qtd_rel, "{{TOT_PLAN}}": tot_plan
    }
    
    st.write("**Dicionário de Tags pronto para injeção:**")
    st.json(mapeamento_tags)
