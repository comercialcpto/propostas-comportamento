"""
Calculadora de logística COMPARTILHADA por todos os módulos.

Retorno: dict no formato de st.session_state.logistica_dados
         {"tipo", "total", "idas", "detalhes"}
"""
import streamlit as st
from ferramentas.utilidades import formatar_moeda, esc_md
from ferramentas import config

OPCOES_LOGISTICA = [
    "1. Sem Logística (Cliente assume / 100% Online)",
    "2. Logística Estimada (Percentual %)",
    "3. Logística Base (Alimentação + Táxi da Base)",
    "4. Logística Completa (Cotações Detalhadas)",
]


def render_logistica(valor_op1, key_prefix, percentual_padrao=30, dias_padrao=5):
    tipo = st.selectbox(
        "Formato de cálculo de deslocamento:",
        OPCOES_LOGISTICA,
        key=f"{key_prefix}_log_tipo",
    )

    total = 0.0
    idas = 0
    detalhes = {}

    if tipo == OPCOES_LOGISTICA[0]:  # Sem Logística
        st.info("A Opção 2 terá o mesmo valor da Opção 1: o cliente arca com os custos de viagem.")

    elif tipo == OPCOES_LOGISTICA[1]:  # Estimada (%)
        perc = st.number_input(
            "Margem Estimada de Logística (%)", min_value=0, max_value=100,
            value=percentual_padrao, key=f"{key_prefix}_log_perc",
        )
        total = valor_op1 * (perc / 100)
        detalhes = {"Percentual Aplicado": f"{perc}% sobre OP1"}
        st.info(esc_md(f"Cálculo: {perc}% sobre o Serviço ({formatar_moeda(valor_op1)}) = {formatar_moeda(total)}"))

    elif tipo == OPCOES_LOGISTICA[2]:  # Base
        c_ida, c_dia = st.columns(2)
        idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1,
                                  key=f"{key_prefix}_log_idas_base")
        dias_ida = c_dia.number_input("Dias por Ida", min_value=1, value=dias_padrao,
                                      key=f"{key_prefix}_log_dias_base")
        custo_taxi = idas * (config.TAXI_BASE_IDA * 2)
        custo_alimentacao = idas * dias_ida * config.CUSTO_ALIMENTACAO_DIA
        total = custo_taxi + custo_alimentacao
        detalhes = {"Táxi (Ida e Volta)": custo_taxi, "Alimentação": custo_alimentacao}
        st.success(esc_md(
            f"**Cálculo Base:** {idas} ida(s) — Táxi: {formatar_moeda(custo_taxi)} + "
            f"{idas * dias_ida} diária(s) de alimentação: {formatar_moeda(custo_alimentacao)} "
            f"= **{formatar_moeda(total)}**"
        ))

    elif tipo == OPCOES_LOGISTICA[3]:  # Completa
        total, idas, detalhes = _logistica_completa(key_prefix, dias_padrao)

    return {"tipo": tipo, "total": total, "idas": idas, "detalhes": detalhes}


def _logistica_completa(key_prefix, dias_padrao):
    c_ida, c_dia = st.columns(2)
    idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1,
                              key=f"{key_prefix}_comp_idas")
    dias_ida = c_dia.number_input("Dias por Ida", min_value=1, value=dias_padrao,
                                  key=f"{key_prefix}_comp_dias")

    st.markdown("##### 🏨 Hospedagem")
    ch1, ch2 = st.columns(2)
    hotel_barato = ch1.number_input("Hotel Mais Barato (R$)", min_value=0.0, step=10.0,
                                    key=f"{key_prefix}_hotel_b")
    hotel_caro = ch2.number_input("Hotel Mais Caro (R$)", min_value=0.0, step=10.0,
                                  key=f"{key_prefix}_hotel_c")
    custo_hotel = ((hotel_barato + hotel_caro) / 2) * dias_ida * idas

    st.markdown("##### ✈️ Passagens Aéreas (+ 10% Taxa)")
    st.caption(f"Fórmula Média: ((Mais Barata + Mais Cara) * {config.AEREO_FATOR}) / {config.AEREO_DIVISOR:.0f}")
    ca1, ca2 = st.columns(2)
    ida_barata = ca1.number_input("Ida: Mais Barata (R$)", min_value=0.0, step=50.0,
                                  key=f"{key_prefix}_ida_b")
    ida_cara = ca2.number_input("Ida: Mais Cara (R$)", min_value=0.0, step=50.0,
                                key=f"{key_prefix}_ida_c")
    ca3, ca4 = st.columns(2)
    volta_barata = ca3.number_input("Volta: Mais Barata (R$)", min_value=0.0, step=50.0,
                                    key=f"{key_prefix}_volta_b")
    volta_cara = ca4.number_input("Volta: Mais Cara (R$)", min_value=0.0, step=50.0,
                                  key=f"{key_prefix}_volta_c")

    media_ida = ((ida_barata + ida_cara) * config.AEREO_FATOR) / config.AEREO_DIVISOR
    media_volta = ((volta_barata + volta_cara) * config.AEREO_FATOR) / config.AEREO_DIVISOR
    custo_aereo_total = (media_ida + media_volta) * config.TAXA_LOGISTICA * idas

    st.markdown("##### 🚗 Carro: No Cliente (Hotel ⇄ Cliente) (+ 10% Taxa)")
    cv1, cv2 = st.columns(2)
    diaria_carro = cv1.number_input("Valor da Diária do Carro (R$)", min_value=0.0, step=10.0,
                                    key=f"{key_prefix}_diaria")
    dist_hotel_cliente = cv2.number_input("Dist. Hotel ⇄ Cliente (Km Total Dia)", min_value=0.0,
                                          step=5.0, key=f"{key_prefix}_dist_h")
    custo_diarias = diaria_carro * dias_ida
    custo_comb_hotel = (dist_hotel_cliente / config.COMBUSTIVEL_KM_POR_LITRO) * config.COMBUSTIVEL_PRECO_LITRO * dias_ida
    custo_carro_cliente = (custo_diarias + custo_comb_hotel) * config.TAXA_LOGISTICA * idas

    st.markdown("##### 🛣️ Carro: Até o Cliente (Aeroporto ⇄ Destino) (+ 10% Taxa)")
    cae1, cae2 = st.columns(2)
    dist_aero_cliente = cae1.number_input("Dist. Aeroporto ⇄ Destino (Km Total Ida e Volta)",
                                          min_value=0.0, step=10.0, key=f"{key_prefix}_dist_a")
    pedagio_aero = cae2.number_input("Pedágios (R$ Totais)", min_value=0.0, step=5.0,
                                     key=f"{key_prefix}_pedagio")
    custo_comb_aero = (dist_aero_cliente / config.COMBUSTIVEL_KM_POR_LITRO) * config.COMBUSTIVEL_PRECO_LITRO
    custo_carro_aero = (pedagio_aero + custo_comb_aero) * config.TAXA_LOGISTICA * idas

    total = custo_hotel + custo_aereo_total + custo_carro_cliente + custo_carro_aero
    detalhes = {
        "Hospedagem": custo_hotel,
        "Passagens Aéreas (c/ taxa)": custo_aereo_total,
        "Carro no Cliente (c/ taxa)": custo_carro_cliente,
        "Carro até Cliente (c/ taxa)": custo_carro_aero,
    }
    st.info(esc_md(
        f"**Resumo da Cotação Detalhada:**\n"
        f"- Hospedagem: {formatar_moeda(custo_hotel)}\n"
        f"- Aéreo (com taxa): {formatar_moeda(custo_aereo_total)}\n"
        f"- Carro no Cliente (com taxa): {formatar_moeda(custo_carro_cliente)}\n"
        f"- Deslocamento Aeroporto (com taxa): {formatar_moeda(custo_carro_aero)}\n"
        f"**Total Logística: {formatar_moeda(total)}**"
    ))
    return total, idas, detalhes
