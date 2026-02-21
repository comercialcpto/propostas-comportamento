import streamlit as st
import datetime

# --- CONFIGURAÇÕES TÉCNICAS (INTELIGÊNCIA DOS 5 LOTES) ---

def calcular_diagnostico_cpto(populacao):
    # Tabela Guarda-Chuva (Lote 4)
    if populacao <= 100: return 108
    elif populacao <= 200: return 128
    elif populacao <= 500: return 144
    elif populacao <= 800: return 160
    else: return 216

def calcular_logistica(ch_total):
    # Faixas de Táxi (Lote 4 - Base Consultor)
    if ch_total <= 250: return 150.0
    elif ch_total <= 500: return 280.0
    else: return 420.0

# --- INTERFACE ---
st.set_page_config(page_title="Emissor CPTO v2.0", layout="wide")
st.title("🚀 Emissor de Propostas Grupo Comportamento - v2.0")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Configurações de Venda")
    valor_hora = st.number_input("Valor Hora Padrão (R$)", value=480.0)
    entidade = st.selectbox("Faturar por:", ["Comportamento (20% imposto)", "Escola (11% imposto)"])
    imposto_rate = 0.20 if "Comportamento" in entidade else 0.11

# --- FORMULÁRIO PRINCIPAL ---
with st.container():
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Identificação do Cliente")
        cliente = st.text_input("Nome da Empresa", placeholder="Ex: Bracell")
        missao = st.text_area("Missão/Valores do Cliente", placeholder="Copie do site ou briefing...")
        dor = st.text_area("Justificativa (A 'Dor')", placeholder="Ex: Baixa percepção de risco na base...")

    with col2:
        st.subheader("2. Definição do Escopo")
        servico = st.selectbox("Selecione o Serviço", [
            "Diagnóstico de Cultura (DCS)",
            "Diagnóstico de Clima (DClima)",
            "Mapeamento de Liderança (MPL)",
            "Riscos Psicossociais (RPS)",
            "Pulse (Diagnóstico Ágil)",
            "EHS Estratégico",
            "Projetos Pontuais / Palestras"
        ])

        # Lógica de inputs específicos por serviço
        ch_final = 0
        if servico in ["Diagnóstico de Cultura (DCS)", "Diagnóstico de Clima (DClima)"]:
            pop = st.number_input("População Total", min_value=1, value=100)
            ch_final = calcular_diagnostico_cpto(pop)
            st.info(f"Critério: Tabela Guarda-Chuva para {pop} pessoas.")

        elif servico == "Mapeamento de Liderança (MPL)":
            n_lideres = st.number_input("Número de Líderes para Mapear", min_value=1, value=10)
            # 6h por líder (2.5 prep + 1 sessão + 2.5 relatório) + 20h base projeto
            ch_final = (n_lideres * 6) + 20 
            st.info(f"Critério: 6h/líder + 20h coordenação.")

        elif servico == "Riscos Psicossociais (RPS)":
            tipo_rps = st.radio("Tipo de RPS", ["Mapeamento (5 meses)", "Gestão Completa (17 meses)"])
            ch_final = 1072 if "Mapeamento" in tipo_rps else 1606
            st.info("Critério: Carga horária fixa conforme Lote 4.")

        elif servico == "Pulse (Diagnóstico Ágil)":
            ch_final = 56
            st.info("Critério: Escopo travado em 56 horas totais.")

        elif servico == "EHS Estratégico":
            ch_final = 112
            st.info("Critério: Jornada de 3 meses conforme modelo.")

        else: # Pontuais
            tipo_p = st.selectbox("Tipo de Evento", ["Palestra Online", "Palestra Presencial", "Imersão Liderança"])
            tabela_p = {"Palestra Online": 30, "Palestra Presencial": 36, "Imersão Liderança": 40}
            ch_final = tabela_p[tipo_p]

# --- PROCESSAMENTO FINANCEIRO (CÁLCULO POR DENTRO) ---
custo_base = ch_final * valor_hora
# Fórmula de Gross-up: Valor / (1 - imposto)
investimento_total = custo_base / (1 - imposto_rate)
taxa_taxi = calcular_logistica(ch_final)

# --- SAÍDA ---
st.markdown("---")
if st.button("🔥 CALCULAR ESTRATÉGIA COMERCIAL"):
    if not cliente:
        st.error("Por favor, digite o nome do cliente.")
    else:
        r1, r2, r3 = st.columns(3)
        r1.metric("Carga Horária Total", f"{ch_final} horas")
        r2.metric("Investimento (Consultoria)", f"R$ {investimento_total:,.2f}")
        r3.metric("Sugestão de Táxi (Ida)", f"R$ {taxa_taxi:.2f}")

        st.subheader("💡 Justificativa Inteligente (Pronta para o Slide)")
        justificativa_texto = f"Alinhado à missão da {cliente} de '{missao}', propomos uma intervenção em {servico} para atuar diretamente sobre {dor}. Utilizaremos a metodologia proprietária da Comportamento para garantir que a segurança deixe de ser um processo e se torne um valor cultural."
        st.success(justificativa_texto)

        st.info(f"📂 Nomenclatura do Arquivo: {datetime.datetime.now().year}_XXX_{cliente}_{servico.split('(')[-1].replace(')','')}")
