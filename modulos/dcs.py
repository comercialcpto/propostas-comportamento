import streamlit as st
import pandas as pd
import math
import datetime
from ferramentas.utilidades import (
    formatar_moeda, valor_por_extenso, calcular_amostra, calcular_amortizacao,
    ir_para, adicionar_unidade, remover_unidade, adicionar_fase, remover_fase
)
from ferramentas.pptx_engine import processar_apresentacao

def render_precificacao():
    st.title("💰 1. Motor de Precificação")
    servico_prec = st.selectbox("Selecione o Serviço para Precificar:", ["Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)"])
    st.session_state.servico_selecionado = servico_prec
    
    if servico_prec == "Diagnóstico (DCS/Clima/DCMA)":
        st.markdown("### 📊 1. Cadastro de População e Unidades")
        
        for i, und in enumerate(st.session_state.unidades_dcs):
            c1, c2, c3, c4 = st.columns([0.4, 0.2, 0.2, 0.2])
            und['nome'] = c1.text_input(f"Nome da Unidade/Área", value=und['nome'], key=f"nome_{und['id']}")
            und['pop_total'] = c2.number_input("População Total", min_value=0, value=und['pop_total'], key=f"pop_{und['id']}")
            und['lideres'] = c3.number_input("Total de Líderes", min_value=0, value=und['lideres'], key=f"lid_{und['id']}")
            
            if len(st.session_state.unidades_dcs) > 1:
                if c4.button("🗑️ Remover", key=f"rem_{und['id']}"):
                    remover_unidade(und['id'])
                    st.rerun()

        st.button("➕ Adicionar Unidade / Área", on_click=adicionar_unidade)
        
        st.markdown("---")
        st.markdown("### ⚙️ 2. Motor Estatístico (Coleta de Dados em Campo)")
        
        total_horas_campo = 0
        total_amostra = 0
        dados_tabela_prova = []
        
        for und in st.session_state.unidades_dcs:
            if und['pop_total'] > 0:
                amostra_total = calcular_amostra(und['pop_total'], margem_erro=0.08, proporcao=0.5)
                amostra_lideres = calcular_amostra(und['lideres'], margem_erro=0.05, proporcao=0.8)
                amostra_operacional = max(0, amostra_total - amostra_lideres)
                
                turmas_hm = math.floor(amostra_lideres / 12)
                entrevistas = amostra_lideres % 12
                turmas_foco = math.ceil(amostra_operacional / 12)
                
                horas_hm = turmas_hm * 2
                horas_foco = turmas_foco * 2
                horas_entrevistas = entrevistas * 1.5
                
                total_horas_unidade = horas_hm + horas_foco + horas_entrevistas
                total_horas_campo += total_horas_unidade
                total_amostra += amostra_total
                
                dados_tabela_prova.append({
                    "Unidade / Área": und['nome'],
                    "Pop. Total": und['pop_total'],
                    "Amostra Base (Het. 8%)": amostra_total,
                    "Pop. Líderes": und['lideres'],
                    "Amostra Líderes (Hom. 5%)": amostra_lideres,
                    "Total Horas Campo": f"{total_horas_unidade} h"
                })
                
                with st.expander(f"📌 Racional de Cálculo: {und['nome']}", expanded=False):
                    st.write(f"**1. Liderança:** Dos {und['lideres']} líderes, a amostra exigida é de **{amostra_lideres} pessoas**.")
                    st.write(f"↳ Formaremos **{turmas_hm} turmas de Hearts & Minds** de 12 pessoas (*{turmas_hm} turmas x 2h = {horas_hm}h*).")
                    st.write(f"↳ O saldo de **{entrevistas} líderes** farão **Entrevistas Individuais** (*{entrevistas} pessoas x 1,5h = {horas_entrevistas}h*).")
                    
                    st.write(f"**2. Base Operacional:** População de {und['pop_total'] - und['lideres']} pessoas, resultando na amostra restante de **{amostra_operacional} pessoas**.")
                    st.write(f"↳ Formaremos **{turmas_foco} Grupos Focais** de até 12 pessoas (*{turmas_foco} turmas x 2h = {horas_foco}h*).")
                    
                    st.success(f"**Total de horas desta unidade:** {horas_hm}h (H&M) + {horas_entrevistas}h (Entrevistas) + {horas_foco}h (Grupos Focais) = **{total_horas_unidade} horas**.")

        if dados_tabela_prova:
            st.markdown("#### 📝 Tabela de Prova Real (Resumo de Amostragem)")
            st.table(pd.DataFrame(dados_tabela_prova))

        st.markdown("---")
        st.markdown("### 📋 3. Plano Detalhado (Etapas Adicionais)")
        st.info("Personalize as etapas fixas e operacionais do projeto. As horas cadastradas aqui serão somadas à coleta de campo.")
        
        taxa_hora = st.number_input("Valor da Taxa Hora Técnica (R$)", min_value=0.0, value=480.0, step=10.0, key="taxa_hora_topo")

        total_horas_fases = 0
        for fase in st.session_state.fases_dcs:
            f1, f2, f3 = st.columns([0.6, 0.2, 0.2])
            fase['nome'] = f1.text_input("Nome da Etapa", value=fase['nome'], key=f"fnome_{fase['id']}")
            fase['horas'] = f2.number_input("Carga Horária (h)", min_value=0, value=fase['horas'], key=f"fhoras_{fase['id']}")
            total_horas_fases += fase['horas']
            if f3.button("🗑️ Remover", key=f"frem_{fase['id']}"):
                remover_fase(fase['id'])
                st.rerun()

        st.button("➕ Adicionar Etapa no Plano", on_click=adicionar_fase)

        valor_parcial_fases = total_horas_fases * taxa_hora
        st.success(f"**Racional do Plano Detalhado:** A soma resultou em **{total_horas_fases} horas cadastradas**. \n\nCálculo de Custos: {total_horas_fases} horas x Taxa de {formatar_moeda(taxa_hora)} = **{formatar_moeda(valor_parcial_fases)}**.")

        st.markdown("---")
        st.markdown("### 💰 4. Precificação Final e Logística")
        
        horas_totais = total_horas_campo + total_horas_fases
        valor_op1 = horas_totais * taxa_hora
        
        st.write("#### Planejamento de Logística (Opção 2)")
        tipo_logistica = st.selectbox("Formato de cálculo de deslocamento:", [
            "1. Sem Logística (100% pelo Cliente)", 
            "2. Logística Base (Alimentação + Táxi da Base)", 
            "3. Logística Completa (Cotações Detalhadas)", 
            "4. Logística Estimada (Percentual %)"
        ])
        
        logistica_total = 0.0
        
        if tipo_logistica == "1. Sem Logística (100% pelo Cliente)":
            st.info("A Opção 2 terá o mesmo valor da Opção 1, pois o cliente arcará com todos os custos de viagem diretamente.")
            logistica_total = 0.0

        elif tipo_logistica == "2. Logística Base (Alimentação + Táxi da Base)":
            c_ida, c_dia = st.columns(2)
            qtd_idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1)
            dias_ida = c_dia.number_input("Dias por Ida (padrão 5 = semana cheia)", min_value=1, value=5)
            
            custo_taxi = qtd_idas * (150 * 2)
            custo_alimentacao = qtd_idas * dias_ida * 120
            logistica_total = custo_taxi + custo_alimentacao
            
            st.success(f"**Cálculo Base:** {qtd_idas} idas (Táxi: R$ {custo_taxi:,.2f}) + {qtd_idas * dias_ida} dias de alimentação (R$ {custo_alimentacao:,.2f}) = **R$ {logistica_total:,.2f}**")

        elif tipo_logistica == "3. Logística Completa (Cotações Detalhadas)":
            c_ida, c_dia = st.columns(2)
            qtd_idas = c_ida.number_input("Quantidade de Idas Presenciais", min_value=1, value=1)
            dias_ida = c_dia.number_input("Dias por Ida (padrão 5)", min_value=1, value=5)
            
            st.markdown("##### 🏨 Hospedagem")
            ch1, ch2 = st.columns(2)
            hotel_barato = ch1.number_input("Diária - Hotel Mais Barato (R$)", min_value=0.0, step=10.0)
            hotel_caro = ch2.number_input("Diária - Hotel Mais Caro (R$)", min_value=0.0, step=10.0)
            media_hotel = (hotel_barato + hotel_caro) / 2
            custo_hotel = media_hotel * dias_ida * qtd_idas
            
            st.markdown("##### ✈️ Passagens Aéreas")
            st.caption("A fórmula = ((Barato + Caro) * 3.2) / 6 será aplicada automaticamente.")
            qtd_trechos = st.number_input("Quantos trechos aéreos diferentes?", min_value=0, value=1)
            custo_aereo_total = 0.0
            for i in range(int(qtd_trechos)):
                ca1, ca2 = st.columns(2)
                aereo_barato = ca1.number_input(f"Trecho {i+1} - Mais Barato (R$)", min_value=0.0, step=50.0, key=f"ab_{i}")
                aereo_caro = ca2.number_input(f"Trecho {i+1} - Mais Caro (R$)", min_value=0.0, step=50.0, key=f"ac_{i}")
                media_aereo = ((aereo_barato + aereo_caro) * 3.2) / 6
                custo_aereo_total += (media_aereo * qtd_idas)

            st.markdown("##### 🚗 Locação de Veículo e Deslocamentos")
            cv1, cv2 = st.columns(2)
            diaria_carro = cv1.number_input("Valor da Diária do Carro (R$)", min_value=0.0, step=10.0)
            custo_carro = diaria_carro * dias_ida * qtd_idas
            
            dist_hotel_cliente = cv2.number_input("Distância Hotel ⇄ Cliente (Km - Total Ida e Volta diária)", min_value=0.0, step=5.0)
            custo_combustivel_hotel = (dist_hotel_cliente / 9.0) * 6.0 * dias_ida * qtd_idas
            
            st.markdown("##### 🛣️ Aeroporto até o Cliente")
            cae1, cae2 = st.columns(2)
            dist_aero_cliente = cae1.number_input("Dist. Aeroporto ⇄ Cliente (Km - Total Ida e Volta da viagem)", min_value=0.0, step=10.0)
            pedagio_aero = cae2.number_input("Valor Pedágios Aeroporto ⇄ Cliente (R$ Total)", min_value=0.0, step=5.0)
            
            custo_combustivel_aero = (dist_aero_cliente / 9.0) * 6.0 * qtd_idas
            custo_pedagio = pedagio_aero * qtd_idas
            
            logistica_total = custo_hotel + custo_aereo_total + custo_carro + custo_combustivel_hotel + custo_combustivel_aero + custo_pedagio
            
            st.info(f"**Resumo Logística Cotação:** Hospedagem (R$ {custo_hotel:,.2f}) + Aéreo (R$ {custo_aereo_total:,.2f}) + Carro (R$ {custo_carro:,.2f}) + Combustível (R$ {(custo_combustivel_hotel+custo_combustivel_aero):,.2f}) + Pedágio (R$ {custo_pedagio:,.2f}) = **R$ {logistica_total:,.2f}**")

        elif tipo_logistica == "4. Logística Estimada (Percentual %)":
            perc_logistica = st.number_input("Margem Estimada de Logística (%)", min_value=0, max_value=100, value=30)
            logistica_total = valor_op1 * (perc_logistica / 100)
            st.info(f"Cálculo: {perc_logistica}% sobre o Serviço Técnico ({formatar_moeda(valor_op1)}) = {formatar_moeda(logistica_total)}")
        
        valor_op2 = valor_op1 + logistica_total
        
        st.session_state.valores_finais["op1"] = valor_op1
        st.session_state.valores_finais["op2"] = valor_op2
        
        st.write(f"**Racional da Precificação Técnica:** {total_horas_campo}h (Campo) + {total_horas_fases}h (Plano Detalhado) = **{horas_totais} horas totais de projeto**.")
        
        c_tot1, c_tot2 = st.columns(2)
        c_tot1.metric("Total OP1 (Serviço Técnico)", formatar_moeda(valor_op1))
        c_tot2.metric("Total OP2 (Com Logística)", formatar_moeda(valor_op2))

        st.write("")
        st.button("Salvar e Avançar para Proposta Técnica ➡️", on_click=ir_para, args=("📝 2. Proposta Técnica",), type="primary")

def render_tecnica():
    st.title("📝 2. Gerador de Proposta Técnica")
    st.caption("As informações inseridas aqui alimentarão automaticamente a Proposta Comercial.")
    
    with st.sidebar:
        st.markdown("---")
        template_upload = st.file_uploader("Suba o template (Técnica)", type="pptx")
    
    campos_vazios = []

    with st.expander("📍 1. Identificação Geral (Compartilhada)", expanded=True):
        c1, c2, c3 = st.columns(3)
        
        srv = st.session_state.get('servico_selecionado', "Diagnóstico (DCS/Clima/DCMA)")
        st.session_state.memoria_geral["servico"] = c1.text_input("Serviço Principal", value=srv, disabled=True)
        
        st.session_state.memoria_geral["cliente"] = c2.text_input("Nome da Empresa ({{CLIENTE}})*", value=st.session_state.memoria_geral.get("cliente", ""))
        st.session_state.memoria_geral["unidade"] = c3.text_input("Unidade ({{UNIDADE}})*", value=st.session_state.memoria_geral.get("unidade", ""))
        
        c4, c5, c6 = st.columns(3)
        st.session_state.memoria_geral["num_prop"] = c4.text_input("Nº da Proposta ({{NUM_PROP}})*", value=st.session_state.memoria_geral.get("num_prop", ""))
        st.session_state.memoria_geral["escopo"] = c5.text_input("Título do Escopo ({{ESCOPO}})*", value=st.session_state.memoria_geral.get("escopo", ""))
        st.session_state.memoria_geral["prazo"] = c6.text_input("Prazo ({{PRAZO}})*", value=st.session_state.memoria_geral.get("prazo", ""))
        
        c7, c8 = st.columns(2)
        st.session_state.memoria_geral["formato"] = c7.selectbox("Formato ({{FORMATO}})*", ["Híbrido", "Presencial", "Online"], index=["Híbrido", "Presencial", "Online"].index(st.session_state.memoria_geral.get("formato", "Híbrido")))
        st.session_state.memoria_geral["idas"] = c8.number_input("Nº de Idas Presenciais ({{IDAS}})", min_value=0, value=st.session_state.memoria_geral.get("idas", 0))
        
        st.session_state.memoria_geral["justificativa"] = st.text_area("Justificativa ({{JUSTIFICATIVA}})*", value=st.session_state.memoria_geral.get("justificativa", ""))
        st.session_state.memoria_geral["objetivo"] = st.text_area("Objetivo ({{OBJETIVO}})*", value=st.session_state.memoria_geral.get("objetivo", ""))

    with st.expander("📅 2. Cronograma de Avanço Inteligente", expanded=True):
        st.info("As fases abaixo foram importadas automaticamente do Plano Detalhado da Precificação.")
        qtd_meses_projeto = st.number_input("Duração total do projeto (meses)", min_value=1, value=12)
        
        atividades_lista = []
        fases_importadas = [{"nome": "Coleta de Dados em Campo"}] + st.session_state.fases_dcs
        
        for i, fase in enumerate(fases_importadas):
            ca, cm = st.columns([0.4, 0.6])
            nome_at = ca.text_input(f"Nome da Fase {i+1}", value=fase['nome'], key=f"tg_{i}")
            meses_at = cm.multiselect("Selecione os meses", list(range(1, int(qtd_meses_projeto) + 1)), key=f"tm_{i}")
            if meses_at:
                atividades_lista.append({"nome": nome_at, "meses": meses_at})

    with st.expander("👥 3. Detalhamento Simplificado do Público", expanded=True):
        st.info("Os totais foram carregados da Precificação. Ajuste caso o projeto demande adicionar terceiros que não estavam no cálculo amostral.")
        
        lideres_calc = sum(u['lideres'] for u in st.session_state.unidades_dcs)
        pop_calc = sum(u['pop_total'] for u in st.session_state.unidades_dcs)
        
        cp1, cp2, cp3 = st.columns(3)
        n_lid_total = cp1.number_input("Total de Líderes (Equipe Própria)", min_value=0, value=int(lideres_calc))
        n_oper = cp2.number_input("Total Operacional (Equipe Própria)", min_value=0, value=int(max(0, pop_calc - lideres_calc)))
        n_terc = cp3.number_input("Terceiros Adicionais", min_value=0, value=0)
        
        n_p_terc = n_lid_total + n_oper + n_terc
        st.metric("Público Alvo Consolidado no Relatório", n_p_terc)

    colA, colB = st.columns(2)
    colA.button("⬅️ Voltar para Precificação", on_click=ir_para, args=("💰 1. Precificação",))
    
    def tentar_gerar_tecnica():
        st.session_state.tentou_gerar = True
        
    colB.button("🚀 VALIDAR E GERAR TÉCNICA", on_click=tentar_gerar_tecnica, type="primary")

    if st.session_state.tentou_gerar:
        if not template_upload:
            st.error("⚠️ Faça o upload do template da Técnica na barra lateral.")
        else:
            mapa = {
                "{{SERVICO}}": st.session_state.memoria_geral["servico"], 
                "{{CLIENTE}}": st.session_state.memoria_geral["cliente"], 
                "{{UNIDADE}}": st.session_state.memoria_geral["unidade"], 
                "{{NUM_PROP}}": st.session_state.memoria_geral["num_prop"], 
                "{{ESCOPO}}": st.session_state.memoria_geral["escopo"],
                "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
                "{{JUSTIFICATIVA}}": st.session_state.memoria_geral["justificativa"], 
                "{{OBJETIVO}}": st.session_state.memoria_geral["objetivo"],
                "{{PUBLICO}}": str(n_p_terc), 
                "{{PRAZO}}": st.session_state.memoria_geral["prazo"], 
                "{{FORMATO}}": st.session_state.memoria_geral["formato"], 
                "{{IDAS}}": str(st.session_state.memoria_geral["idas"]),
                "{{N_LID}}": str(n_lid_total), 
                "{{N_OPER}}": str(n_oper), 
                "{{N_PTERC}}": str(n_p_terc)
            }
            with st.spinner("Construindo arquivo..."):
                st.session_state.pptx_gerado = processar_apresentacao(template_upload, mapa, atividades_lista, "Técnica", None, qtd_meses_projeto)
                st.session_state.nome_arquivo = f"Tecnica_{st.session_state.memoria_geral['cliente']}.pptx"
            st.success("Técnica gerada com sucesso!")
            st.session_state.tentou_gerar = False

    if st.session_state.pptx_gerado and st.session_state.nome_arquivo.startswith("Tecnica"):
        st.download_button("⬇️ Baixar PPTX Técnico", data=st.session_state.pptx_gerado, file_name=st.session_state.nome_arquivo)
        st.write("")
        st.button("Avançar para Proposta Comercial ➡️", on_click=ir_para, args=("📈 3. Proposta Comercial",))

def render_comercial():
    st.title("📈 3. Gerador de Proposta Comercial")
    st.caption("A mágica acontece aqui: não é preciso digitar mais nada além das parcelas.")
    
    with st.sidebar:
        st.markdown("---")
        template_upload = st.file_uploader("Suba o template (Comercial)", type="pptx")

    st.success("✅ Informações Gerais (Cliente, Justificativa, Prazos) carregadas da Técnica.")
    
    with st.expander("💰 Configuração de Parcelamento", expanded=True):
        st.info("Os valores totais foram puxados diretamente do Motor de Precificação.")
        v_op1 = st.session_state.valores_finais["op1"]
        v_op2 = st.session_state.valores_finais["op2"]
        
        st.markdown(f"**Total Investimento Técnico (OP1):** {formatar_moeda(v_op1)}")
        st.markdown(f"**Total Turnkey com Logística (OP2):** {formatar_moeda(v_op2)}")
        
        st.write("---")
        qtd_parcelas = st.number_input("Quantidade de Parcelas ({{QTD_PARCELAS}})", min_value=1, value=12)

    colA, colB = st.columns(2)
    colA.button("⬅️ Voltar para Técnica", on_click=ir_para, args=("📝 2. Proposta Técnica",))
    
    def tentar_gerar_comercial():
        st.session_state.tentou_gerar = True
        
    colB.button("🚀 VALIDAR E GERAR COMERCIAL", on_click=tentar_gerar_comercial, type="primary")

    if st.session_state.tentou_gerar:
        if not template_upload:
            st.error("⚠️ Faça o upload do template da Comercial na barra lateral.")
        else:
            mapa_comercial = {
                "{{SERVICO}}": st.session_state.memoria_geral["servico"], 
                "{{CLIENTE}}": st.session_state.memoria_geral["cliente"], 
                "{{UNIDADE}}": st.session_state.memoria_geral["unidade"], 
                "{{NUM_PROP}}": st.session_state.memoria_geral["num_prop"], 
                "{{ESCOPO}}": st.session_state.memoria_geral["escopo"],
                "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
                "{{JUSTIFICATIVA}}": st.session_state.memoria_geral["justificativa"], 
                "{{OBJETIVO}}": st.session_state.memoria_geral["objetivo"],
                "{{PRAZO}}": st.session_state.memoria_geral["prazo"], 
                "{{FORMATO}}": st.session_state.memoria_geral["formato"], 
                "{{IDAS}}": str(st.session_state.memoria_geral["idas"]),
                "{{VALOR_OP1}}": formatar_moeda(v_op1),
                "{{VALOR_OP2}}": formatar_moeda(v_op2),
                "{{VALOR_OP1_EXT}}": valor_por_extenso(v_op1),
                "{{VALOR_OP2_EXT}}": valor_por_extenso(v_op2),
                "{{QTD_PARCELAS}}": str(qtd_parcelas),
                "{{VLR1_PARCELAS}}": formatar_moeda(v_op1/qtd_parcelas) if qtd_parcelas > 0 else "0",
                "{{VLR2_PARCELAS}}": formatar_moeda(v_op2/qtd_parcelas) if qtd_parcelas > 0 else "0"
            }
            
            acoes_auto = [{"nome": "Diagnóstico de Cultura e Intervenção", "v1": v_op1, "v2": v_op2}]
            dist_parcelas = calcular_amortizacao(qtd_parcelas)
            
            dados_financeiros = {
                'acoes': acoes_auto, 
                'total_op1': v_op1, 
                'total_op2': v_op2,
                'parcelas': dist_parcelas
            }
            
            with st.spinner("Construindo arquivo..."):
                st.session_state.pptx_gerado = processar_apresentacao(template_upload, mapa_comercial, [], "Comercial", dados_financeiros, 12)
                st.session_state.nome_arquivo = f"Comercial_{st.session_state.memoria_geral['cliente']}.pptx"
            st.success("Comercial gerada com sucesso!")
            st.session_state.tentou_gerar = False

    if st.session_state.pptx_gerado and st.session_state.nome_arquivo.startswith("Comercial"):
        st.download_button("⬇️ Baixar PPTX Comercial", data=st.session_state.pptx_gerado, file_name=st.session_state.nome_arquivo)
