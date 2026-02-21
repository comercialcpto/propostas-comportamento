import streamlit as st
import datetime

# --- INTELIGÊNCIA DE NEGÓCIO ---
def calcular_diagnostico_cpto(populacao):
    if populacao <= 100: return 108
    elif populacao <= 200: return 128
    elif populacao <= 500: return 144
    elif populacao <= 800: return 160
    else: return 216

def calcular_taxa_taxi(ch_total):
    if ch_total <= 250: return 150.0
    elif ch_total <= 500: return 280.0
    else: return 420.0

# --- INTERFACE ---
st.set_page_config(page_title="Emissor CPTO v3.0", layout="wide")
st.title("🚀 Emissor de Propostas Grupo Comportamento - v3.0")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Parâmetros Financeiros")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Faturamento:", ["Comportamento (20%)", "Escola (11%)"])
    imposto_rate = 0.20 if "20%" in entidade else 0.11

# --- FORMULÁRIO ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Identificação e Justificativa")
    cliente = st.text_input("Nome da Empresa")
    missao = st.text_area("Missão do Cliente")
    dor = st.text_area("Dor/Cenário do Cliente")
    servico = st.selectbox("Serviço", ["Diagnóstico (DCS/Clima)", "Liderança (MPL)", "RPS", "Pulse", "EHS Estratégico", "Pontuais"])

with col2:
    st.subheader("2. Dimensionamento e Logística")
    ch_total = 0
    if servico == "Diagnóstico (DCS/Clima)":
        pop = st.number_input("População", min_value=1, value=100)
        ch_total = calcular_diagnostico_cpto(pop)
    elif servico == "Liderança (MPL)":
        n_lideres = st.number_input("Nº de Líderes", value=10)
        ch_total = (n_lideres * 6) + 20
    elif servico == "RPS":
        ch_total = st.radio("Escopo RPS", [1072, 1606])
    elif servico == "Pulse": ch_total = 56
    elif servico == "EHS Estratégico": ch_total = 112
    else: ch_total = st.number_input("CH Manual", value=16)

    st.markdown("---")
    st.write("**Custos de Viagem por Ida:**")
    n_idas = st.number_input("Número de Idas Presenciais", min_value=0, value=1)
    val_aereo = st.number_input("Média Aéreo (R$)", value=1200.0)
    val_hotel = st.number_input("Diária Hotel (R$)", value=350.0)
    n_pernoites = st.number_input("Nº de Pernoites por Ida", value=4)

# --- MOTOR DE CÁLCULO ---
# 1. Consultoria
custo_consultoria_bruto = ch_total * valor_hora
investimento_cpto = custo_consultoria_bruto / (1 - imposto_rate)

# 2. Logística (Regra do Lote 4)
taxa_taxi = calcular_taxa_taxi(ch_total)
# Cada ida tem: 2 taxis (casa/aero) + Alimentação (n_pernoites+1)
custo_taxi_base = (taxa_taxi * 2) * n_idas
custo_alimentacao = (120.0 * (n_pernoites + 1)) * n_idas

# Opção 2: Tudo incluso (Aéreo + Hotel + Taxi + Almoço)
custo_viagem_total = ((val_aereo) + (val_hotel * n_pernoites) + (taxa_taxi * 2) + (120.0 * (n_pernoites + 1))) * n_idas
logistica_inclusa = custo_viagem_total / (1 - imposto_rate)

# --- PAINEL DE RESULTADOS ---
if st.button("🔥 GERAR ESTRATÉGIA COMERCIAL"):
    st.markdown("---")
    res1, res2 = st.columns(2)
    
    with res1:
        st.success("### Opção 1: Reembolso")
        st.write("Cliente paga aéreos/hotéis à parte ou via Nota de Débito.")
        st.metric("Investimento", f"R$ {investimento_cpto:,.2f}")
        st.caption(f"Incluso: Consultoria ({ch_total}h) + Táxi Base + Alimentação.")

    with res2:
        st.info("### Opção 2: Logística Inclusa")
        st.write("Valor global com todas as despesas embutidas.")
        st.metric("Investimento Total", f"R$ {(investimento_cpto + logistica_inclusa):,.2f}")
        st.caption(f"Incluso: Consultoria + {n_idas} idas (Aéreo/Hotel/Alimentação/Táxis).")

    st.markdown("---")
    st.subheader("📂 Documentação")
    st.code(f"Nomenclatura: 2026_XXX_{cliente.replace(' ','_')}_{servico}")
    
    st.subheader("💡 Justificativa Técnica")
    st.write(f"Para apoiar a {cliente} em sua missão de '{missao}', focaremos em {dor} através de {ch_total}h de consultoria especializada.")
