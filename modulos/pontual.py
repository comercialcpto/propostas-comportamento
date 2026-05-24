import streamlit as st
import pandas as pd
from ferramentas.utilidades import formatar_moeda, ir_para

# MATRIZ DE PREÇOS OFICIAL
CATALOGO_PONTUAL = {
    "Consultor": {
        "Até 2 horas": {"Presencial": 6999.00, "Online": 4990.00},
        "Até 4 horas": {"Presencial": 13500.00, "Online": 13500.00},
        "De 5 a 8 horas": {"Presencial": 27000.00, "Online": 22500.00},
        "Acima de 8 horas": {"Presencial": 37800.00, "Online": 31500.00}
    },
    "Diretoria": {
        "Até 4 horas": {"Presencial": 22680.00, "Online": 18900.00},
        "De 5 a 8 horas": {"Presencial": 37800.00, "Online": 31500.00},
        "Acima de 8 horas": {"Presencial": 52920.00, "Online": 44100.00}
    },
    "Produto Digital (E-book)": {
        "OAC": {"Digital": 39.90},
        "DDS": {"Digital": 39.90},
        "IAO": {"Digital": 39.90},
        "PGC": {"Digital": 39.90},
        "Pack Completo": {"Digital": 119.70}
    },
    "Produto Digital (E-book + Mentoria)": {
        "OAC": {"Digital": 99.90},
        "DDS": {"Digital": 99.90},
        "IAO": {"Digital": 99.90},
        "PGC": {"Digital": 99.90},
        "Pack Completo": {"Digital": 299.70}
    },
    "Produto Digital (E-book - EV)": {
        "Segurança Psicológica": {"Digital": 39.90},
        "Gestão das Emoções": {"Digital": 39.90},
        "Inovação e Adaptabilidade": {"Digital": 39.90},
        "Comunicação Efetiva": {"Digital": 39.90}
    }
}

def render_precificacao():
    
    # --- 1. POPULAÇÃO E UNIDADES ---
    st.markdown("### 👥 1. Cadastro de População e Unidades")
    st.info("Identifique os públicos para direcionar as palestras/workshops. Não há cálculo amostral neste módulo.")
    
    for i, und in enumerate(st.session_state.unidades_pontual):
        c1, c2, c3, c4 = st.columns([0.4, 0.2, 0.2, 0.2])
        und['nome'] = c1.text_input(f"Nome do Público/Área", value=und['nome'], key=f"p_nome_{und['id']}")
        und['pop_total'] = c2.number_input("População Total", min_value=0, value=und['pop_total'], key=f"p_pop_{und['id']}")
        und['lideres'] = c3.number_input("Destes, Líderes", min_value=0, value=und['lideres'], key=f"p_lid_{und['id']}")
        
        if len(st.session_state.unidades_pontual) > 1:
            if c4.button("🗑️ Remover", key=f"p_rem_{und['id']}"):
                st.session_state.unidades_pontual = [u for u in st.session_state.unidades_pontual if u['id'] != und['id']]
                st.rerun()

    def add_und_pontual():
        novo_id = len(st.session_state.unidades_pontual)
        st.session_state.unidades_pontual.append({"id": novo_id, "nome": f"Novo Público {novo_id+1}", "pop_total": 0, "lideres": 0})
        
    st.button("➕ Adicionar Público / Área", on_click=add_und_pontual)
    
    st.markdown("---")

    # --- 2. SELEÇÃO DA MATRIZ (O PREÇO VEM DAQUI) ---
    st.markdown("### 🛒 2. Seleção de Soluções (Composição de Preço)")
    st.info("O investimento financeiro da proposta é calculado exclusivamente pelos itens adicionados abaixo.")

    with st.container():
        c1, c2, c3, c4 = st.columns([0.3, 0.3, 0.2, 0.2])
        item_sel = c1.selectbox("Profissional / Produto", list(CATALOGO_PONTUAL.keys()))
        categoria_sel = c2.selectbox("Duração / Tema", list(CATALOGO_PONTUAL[item_sel].keys()))
        formato_sel = c3.selectbox("Formato", list(CATALOGO_PONTUAL[item_sel][categoria_sel].keys()))
        qtd_sel = c4.number_input("Quantidade", min_value=1, value=1)

    valor_unitario = CATALOGO_PONTUAL[item_sel][categoria_sel][formato_sel]
    valor_total_item = valor_unitario * qtd_sel
    
    col_add, col_vlr = st.columns([0.2, 0.8])
    if col_add.button("➕ Adicionar ao Escopo", type="primary"):
        st.session_state.carrinho_pontual.append({
            "id": len(st.session_state.carrinho_pontual),
            "Item": item_sel,
            "Detalhe": categoria_sel,
            "Formato": formato_sel,
            "Qtd": qtd_sel,
            "Subtotal": valor_total_item
        })
        st.rerun()
    
    col_vlr.markdown(f"**Valor do item selecionado:** {qtd_sel}x de {formatar_moeda(valor_unitario)} = **{formatar_moeda(valor_total_item)}**")

    # Exibe o Carrinho
    if len(st.session_state.carrinho_pontual) == 0:
        st.warning("Nenhum item financeiro adicionado à proposta.")
        valor_op1 = 0.0
    else:
        for idx, item in enumerate(st.session_state.carrinho_pontual):
            rc1, rc2, rc3 = st.columns([0.7, 0.2, 0.1])
            rc1.write(f"**{item['Qtd']}x {item['Item']}** ({item['Detalhe']} - {item['Formato']})")
            rc2.write(f"{formatar_moeda(item['Subtotal'])}")
            if rc3.button("🗑️", key=f"del_pontual_{idx}"):
                st.session_state.carrinho_pontual.pop(idx)
                st.rerun()
        valor_op1 = sum(item["Subtotal"] for item in st.session_state.carrinho_pontual)
        st.success(f"**Total Financeiro do Escopo (OP1):** {formatar_moeda(valor_op1)}")

    st.markdown("---")

    # --- 3. PLANO DETALHADO (AS FASES PADRÃO) ---
    st.markdown("### 📋 3. Plano Detalhado do Projeto")
    st.info("As etapas de Abertura, Desenvolvimento e Análise Crítica já estão bonificadas dentro do valor da Matriz selecionada acima. Ajuste a carga horária se necessário para refletir o esforço real no relatório.")

    total_horas_fases = 0
    for fase in st.session_state.fases_pontual:
        f1, f2, f3 = st.columns([0.6, 0.2, 0.2])
        fase['nome'] = f1.text_input("Nome da Etapa", value=fase['nome'], key=f"fpnome_{fase['id']}")
        fase['horas'] = f2.number_input("Carga Horária (h)", min_value=0, value=fase['horas'], key=f"fphoras_{fase['id']}")
        total_horas_fases += fase['horas']
        
        if f3.button("🗑️ Remover", key=f"fprem_{fase['id']}"):
            st.session_state.fases_pontual = [f for f in st.session_state.fases_pontual if f['id'] != fase['id']]
            st.rerun()

    def add_fase_pontual():
        novo_id = max([f['id'] for f in st.session_state.fases_pontual], default=-1) + 1
        st.session_state.fases_pontual.append({"id": novo_id, "nome": "Nova Etapa", "horas": 0})

    st.button("➕ Adicionar Etapa ao Cronograma", on_click=add_fase_pontual)
    st.markdown(f"**Total de Carga Horária do Projeto:** {total_horas_fases} horas")

    st.markdown("---")

    # --- 4. LOGÍSTICA ---
    st.markdown("### ✈️ 4. Planejamento de Logística")
    
    tipo_logistica = st.selectbox("Formato de cálculo de deslocamento:", [
        "1. Sem Logística (100% Online ou Cliente assume)", 
        "2. Logística Estimada (Percentual %)",
        "3. Logística Base (Alimentação + Táxi da Base)",
        "4. Logística Completa (Cotações Detalhadas)"
    ])
    
    logistica_total = 0.0
    
    if tipo_logistica == "1. Sem Logística (100% Online ou Cliente assume)":
        logistica_total = 0.0

    elif tipo_logistica == "2. Logística Estimada (Percentual %)":
        perc_logistica = st.number_input("Margem Estimada de Logística (%)", min_value=0, max_value=100, value=15)
        logistica_total = valor_op1 * (perc_logistica / 100)
        st.info(f"Cálculo: {perc_logistica}% sobre o Serviço ({formatar_moeda(valor_op1)}) = {formatar_moeda(logistica_total)}")

    elif tipo_logistica == "3. Logística Base (Alimentação + Táxi da Base)":
        c_ida, c_dia = st.columns(2)
        qtd_idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1)
        dias_ida = c_dia.number_input("Dias por Ida (Palestras geralmente 1 ou 2)", min_value=1, value=1)
        custo_taxi = qtd_idas * (150 * 2)
        custo_alimentacao = qtd_idas * dias_ida * 120
        logistica_total = custo_taxi + custo_alimentacao
        st.success(f"Táxi: R$ {custo_taxi:,.2f} + Alimentação: R$ {custo_alimentacao:,.2f} = **R$ {logistica_total:,.2f}**")

    elif tipo_logistica == "4. Logística Completa (Cotações Detalhadas)":
        c_ida, c_dia = st.columns(2)
        qtd_idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1)
        dias_ida = c_dia.number_input("Dias por Ida", min_value=1, value=1)
        
        st.markdown("##### 🏨 Hospedagem")
        ch1, ch2 = st.columns(2)
        hotel_barato = ch1.number_input("Hotel Mais Barato (R$)", min_value=0.0, step=10.0)
        hotel_caro = ch2.number_input("Hotel Mais Caro (R$)", min_value=0.0, step=10.0)
        custo_hotel = ((hotel_barato + hotel_caro) / 2) * dias_ida * qtd_idas
        
        st.markdown("##### ✈️ Passagens Aéreas (+ 10% Taxa)")
        st.caption("Fórmula Média: ((Mais Barata + Mais Cara) * 3.2) / 6")
        ca1, ca2 = st.columns(2)
        ida_barata = ca1.number_input("Ida: Mais Barata (R$)", min_value=0.0, step=50.0, key="ida_b_pont")
        ida_cara = ca2.number_input("Ida: Mais Cara (R$)", min_value=0.0, step=50.0, key="ida_c_pont")
        
        ca3, ca4 = st.columns(2)
        volta_barata = ca3.number_input("Volta: Mais Barata (R$)", min_value=0.0, step=50.0, key="volta_b_pont")
        volta_cara = ca4.number_input("Volta: Mais Cara (R$)", min_value=0.0, step=50.0, key="volta_c_pont")
        
        media_ida = ((ida_barata + ida_cara) * 3.2) / 6
        media_volta = ((volta_barata + volta_cara) * 3.2) / 6
        custo_aereo_total = (media_ida + media_volta) * 1.10 * qtd_idas

        st.markdown("##### 🚗 Carro: No Cliente (Hotel ⇄ Cliente) (+ 10% Taxa)")
        cv1, cv2 = st.columns(2)
        diaria_carro = cv1.number_input("Valor da Diária (R$)", min_value=0.0, step=10.0, key="diaria_c_pont")
        dist_hotel_cliente = cv2.number_input("Dist. Hotel ⇄ Cliente (Km Total Dia)", min_value=0.0, step=5.0, key="dist_h_pont")
        
        custo_diarias = diaria_carro * dias_ida
        custo_comb_hotel = (dist_hotel_cliente / 9.0) * 6.0 * dias_ida
        custo_carro_cliente = (custo_diarias + custo_comb_hotel) * 1.10 * qtd_idas
        
        st.markdown("##### 🛣️ Carro: Até o Cliente (Aeroporto ⇄ Destino) (+ 10% Taxa)")
        cae1, cae2 = st.columns(2)
        dist_aero_cliente = cae1.number_input("Dist. Aeroporto ⇄ Destino (Km Total Ida e Volta)", min_value=0.0, step=10.0, key="dist_a_pont")
        pedagio_aero = cae2.number_input("Pedágios (R$ Totais)", min_value=0.0, step=5.0, key="ped_a_pont")
        
        custo_comb_aero = (dist_aero_cliente / 9.0) * 6.0
        custo_carro_aero = (pedagio_aero + custo_comb_aero) * 1.10 * qtd_idas
        
        logistica_total = custo_hotel + custo_aereo_total + custo_carro_cliente + custo_carro_aero
        
        st.info(f"**Resumo da Cotação Detalhada:** \n"
                f"- Hospedagem: {formatar_moeda(custo_hotel)} \n"
                f"- Aéreo (com taxa): {formatar_moeda(custo_aereo_total)} \n"
                f"- Carro no Cliente (com taxa): {formatar_moeda(custo_carro_cliente)} \n"
                f"- Deslocamento Aeroporto (com taxa): {formatar_moeda(custo_carro_aero)} \n"
                f"**Total Logística: {formatar_moeda(logistica_total)}**")

    valor_op2 = valor_op1 + logistica_total
    
    st.session_state.valores_finais = {"op1": valor_op1, "op2": valor_op2}

    st.markdown("---")
    st.markdown("### 💰 5. Resumo da Precificação")
    
    c_tot1, c_tot2 = st.columns(2)
    c_tot1.metric("Total OP1 (Baseado na Matriz)", formatar_moeda(valor_op1))
    c_tot2.metric("Total OP2 (Turnkey com Logística)", formatar_moeda(valor_op2))

    st.write("")
    st.button("Salvar e Avançar para Técnica ➡️", on_click=ir_para, args=("📝 2. Proposta Técnica",), type="primary")

def render_tecnica():
    st.title("📝 2. Proposta Técnica (Pontual)")
    st.info("A geração de PPTX para propostas pontuais será implementada na próxima etapa, usando o mesmo modelo do DCS.")
    st.button("Avançar para Proposta Comercial ➡️", on_click=ir_para, args=("📈 3. Proposta Comercial",))

def render_comercial():
    st.title("📈 3. Proposta Comercial (Pontual)")
    st.info("A geração de PPTX comercial para propostas pontuais será implementada na próxima etapa.")
