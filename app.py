import streamlit as st
import pandas as pd
import datetime

# --- INTELIGÊNCIA CPTO (LOTE 4 - FINANCEIRO) ---
def calcular_ch_diagnostico(populacao):
    if populacao <= 100: return 108
    elif populacao <= 200: return 128
    elif populacao <= 500: return 144
    elif populacao <= 800: return 160
    else: return 216

TABELA_PONTUAIS = {
    "Palestra Online (Consultor)": 30,
    "Palestra Presencial (Consultor)": 36,
    "Palestra Online (Diretoria)": 42,
    "Palestra Presencial (Diretoria)": 48,
    "Imersão Média Liderança": 40,
}

# --- INTERFACE ---
st.set_page_config(page_title="Emissor CPTO", layout="centered")
st.title("🎯 Emissor de Propostas Grupo Comportamento")

with st.form("proposta_form"):
    st.subheader("1. Identificação")
    cliente = st.text_input("Nome do Cliente")
    missao = st.text_area("Palavras-chave da Missão do Cliente")
    dor = st.text_area("Dor/Justificativa do Cliente")
    
    st.subheader("2. Escopo")
    categoria = st.radio("Categoria", ["Diagnóstico (DCS/RPS)", "Projetos Pontuais"])
    
    if categoria == "Diagnóstico (DCS/RPS)":
        n_pessoas = st.number_input("População Total", min_value=1, value=100)
        ch_total = calcular_ch_diagnostico(n_pessoas)
    else:
        tipo = st.selectbox("Tipo de Evento", list(TABELA_PONTUAIS.keys()))
        ch_total = TABELA_PONTUAIS[tipo]

    logistica = st.selectbox("Logística", ["Incluso na Proposta", "Reembolso/Nota de Débito"])
    
    submitted = st.form_submit_button("Calcular Proposta")

if submitted:
    # Lógica Lote 4
    valor_total = (ch_total * 480) * 1.25  # Gross-up de 20%
    
    taxa_taxi = 150.0
    if ch_total > 250: taxa_taxi = 280.0
    if ch_total > 500: taxa_taxi = 420.0

    st.success("✅ Cálculos Realizados!")
    
    c1, c2 = st.columns(2)
    c1.metric("Carga Horária", f"{ch_total}h")
    c2.metric("Investimento Total", f"R$ {valor_total:,.2f}")
    
    st.warning(f"Sugestão de Táxi (Base Consultor): R$ {taxa_taxi:.2f} por deslocamento")
    
    st.subheader("📝 Justificativa Técnica Sugerida")
    justificativa = f"Para que a {cliente} possa '{missao}', propomos uma jornada focada em {dor}, utilizando a metodologia Hearts & Minds para garantir a evolução da maturidade preventiva."
    st.write(justificativa)
    
    st.info(f"Nomenclatura Recomendada: {datetime.datetime.now().year}_XXX_{cliente}_Proposta")
