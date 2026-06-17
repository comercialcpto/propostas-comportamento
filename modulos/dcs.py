import streamlit as st
import pandas as pd
import math
import datetime
from ferramentas.utilidades import (
    formatar_moeda, valor_por_extenso, calcular_amostra, calcular_amortizacao,
    ir_para, adicionar_unidade, remover_unidade, adicionar_fase, remover_fase, esc_md
)
from ferramentas.pptx_engine import processar_apresentacao
from ferramentas.logistica import render_logistica
from ferramentas import config


def render_precificacao():
    # O router já garante que só chegamos aqui no serviço de Diagnóstico.
    st.markdown("### 📊 1. Cadastro de População e Unidades")
    st.caption(
        "Para cada unidade: população e líderes (entram no cálculo amostral) e as atividades "
        "manuais — OAC, Visitas Técnicas e Aprofundamento — que somam horas presenciais de campo."
    )

    for i, und in enumerate(st.session_state.unidades_dcs):
        c1, c2, c3, c4 = st.columns([0.4, 0.2, 0.2, 0.2])
        und['nome'] = c1.text_input("Nome da Unidade/Área", value=und['nome'], key=f"nome_{und['id']}")
        und['pop_total'] = c2.number_input("População Total", min_value=0, value=und['pop_total'], key=f"pop_{und['id']}")
        und['lideres'] = c3.number_input("Total de Líderes", min_value=0, value=und['lideres'], key=f"lid_{und['id']}")

        if len(st.session_state.unidades_dcs) > 1:
            if c4.button("🗑️ Remover", key=f"rem_{und['id']}"):
                remover_unidade(und['id'])
                st.rerun()

        # Atividades manuais desta unidade (somam às horas de campo, todas presenciais).
        # .get() protege unidades antigas em sessão que ainda não tenham as chaves novas.
        m1, m2, m3 = st.columns(3)
        und['oac'] = m1.number_input(
            f"OAC (× {config.HORAS_POR_OAC}h)", min_value=0,
            value=und.get('oac', 0), key=f"oac_{und['id']}"
        )
        und['visitas'] = m2.number_input(
            f"Visitas Técnicas (× {config.HORAS_POR_VISITA}h)", min_value=0,
            value=und.get('visitas', 0), key=f"vis_{und['id']}"
        )
        und['aprofundamento'] = m3.number_input(
            f"Atividades de Aprofundamento (× {config.HORAS_POR_APROFUNDAMENTO}h)", min_value=0,
            value=und.get('aprofundamento', 0), key=f"apr_{und['id']}"
        )

    st.button("➕ Adicionar Unidade / Área", on_click=adicionar_unidade)

    st.markdown("---")
    st.markdown("### ⚙️ 2. Motor Estatístico (Coleta de Dados em Campo)")

    total_horas_campo = 0
    total_amostra = 0
    dados_tabela_prova = []

    for und in st.session_state.unidades_dcs:
        if und['pop_total'] > 0:
            amostra_total = calcular_amostra(und['pop_total'],
                                             margem_erro=config.MARGEM_ERRO_GERAL,
                                             proporcao=config.PROPORCAO_GERAL)
            amostra_lideres = calcular_amostra(und['lideres'],
                                               margem_erro=config.MARGEM_ERRO_LIDERES,
                                               proporcao=config.PROPORCAO_LIDERES)
            amostra_operacional = max(0, amostra_total - amostra_lideres)

            turmas_hm = math.floor(amostra_lideres / config.TAMANHO_TURMA)
            entrevistas = amostra_lideres % config.TAMANHO_TURMA
            turmas_foco = math.ceil(amostra_operacional / config.TAMANHO_TURMA)

            horas_hm = turmas_hm * config.HORAS_POR_TURMA
            horas_foco = turmas_foco * config.HORAS_POR_TURMA
            horas_entrevistas = entrevistas * config.HORAS_POR_ENTREVISTA

            # Atividades manuais desta unidade (cadastradas na Seção 1).
            qtd_oac = und.get('oac', 0)
            qtd_visitas = und.get('visitas', 0)
            qtd_aprof = und.get('aprofundamento', 0)

            horas_oac = qtd_oac * config.HORAS_POR_OAC
            horas_visitas = qtd_visitas * config.HORAS_POR_VISITA
            horas_aprof = qtd_aprof * config.HORAS_POR_APROFUNDAMENTO
            horas_manuais = horas_oac + horas_visitas + horas_aprof

            total_horas_unidade = horas_hm + horas_foco + horas_entrevistas + horas_manuais
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
                st.write(f"↳ Formaremos **{turmas_hm} turmas de Hearts & Minds** de {config.TAMANHO_TURMA} pessoas (*{turmas_hm} turmas x {config.HORAS_POR_TURMA}h = {horas_hm}h*).")
                st.write(f"↳ O saldo de **{entrevistas} líderes** farão **Entrevistas Individuais** (*{entrevistas} pessoas x {config.HORAS_POR_ENTREVISTA}h = {horas_entrevistas}h*).")

                st.write(f"**2. Base Operacional:** População de {und['pop_total'] - und['lideres']} pessoas, resultando na amostra restante de **{amostra_operacional} pessoas**.")
                st.write(f"↳ Formaremos **{turmas_foco} Grupos Focais** de até {config.TAMANHO_TURMA} pessoas (*{turmas_foco} turmas x {config.HORAS_POR_TURMA}h = {horas_foco}h*).")

                if horas_manuais > 0:
                    st.write(
                        f"**3. Atividades Manuais (presenciais):** "
                        f"OAC (*{qtd_oac} × {config.HORAS_POR_OAC}h = {horas_oac}h*) · "
                        f"Visitas Técnicas (*{qtd_visitas} × {config.HORAS_POR_VISITA}h = {horas_visitas}h*) · "
                        f"Aprofundamento (*{qtd_aprof} × {config.HORAS_POR_APROFUNDAMENTO}h = {horas_aprof}h*) "
                        f"= **{horas_manuais}h**."
                    )

                st.success(
                    f"**Total de horas desta unidade:** {horas_hm}h (H&M) + {horas_entrevistas}h (Entrevistas) + "
                    f"{horas_foco}h (Grupos Focais) + {horas_manuais}h (Atividades Manuais) = "
                    f"**{total_horas_unidade} horas**."
                )

    if dados_tabela_prova:
        st.markdown("#### 📝 Tabela de Prova Real (Resumo de Amostragem)")
        st.table(pd.DataFrame(dados_tabela_prova))

    # Contador de modalidade (Presencial x Online). Fica logo abaixo da Prova Real,
    # mas é preenchido após a Seção 3 — precisa do split das etapas do Plano Detalhado.
    placeholder_modalidade = st.container()

    st.markdown("---")
    st.markdown("### 📋 3. Plano Detalhado (Etapas Adicionais)")
    st.info(
        "Personalize as etapas fixas e operacionais do projeto. As horas cadastradas aqui serão "
        "somadas à coleta de campo. Marque **Presencial?** nas etapas que acontecem presencialmente "
        "— as demais entram como horas online."
    )

    taxa_hora = st.number_input("Valor da Hora Técnica (R$)", min_value=0.0,
                                value=config.TAXA_HORA_PADRAO, step=10.0, key="taxa_hora_topo")

    total_horas_fases = 0
    horas_fases_presencial = 0
    horas_fases_online = 0
    for fase in st.session_state.fases_dcs:
        f1, f2, f3, f4 = st.columns([0.46, 0.18, 0.18, 0.18])
        fase['nome'] = f1.text_input("Nome da Etapa", value=fase['nome'], key=f"fnome_{fase['id']}")
        fase['horas'] = f2.number_input("Carga Horária (h)", min_value=0, value=fase['horas'], key=f"fhoras_{fase['id']}")
        fase['presencial'] = f3.checkbox("Presencial?", value=fase.get('presencial', False), key=f"fpres_{fase['id']}")

        total_horas_fases += fase['horas']
        if fase['presencial']:
            horas_fases_presencial += fase['horas']
        else:
            horas_fases_online += fase['horas']

        if f4.button("🗑️ Remover", key=f"frem_{fase['id']}"):
            remover_fase(fase['id'])
            st.rerun()

    st.button("➕ Adicionar Etapa no Plano", on_click=adicionar_fase)

    valor_parcial_fases = total_horas_fases * taxa_hora
    st.success(esc_md(f"**Racional do Plano Detalhado:** A soma resultou em **{total_horas_fases} horas cadastradas**. \n\nCálculo de Custos: {total_horas_fases} horas x Hora de {formatar_moeda(taxa_hora)} = **{formatar_moeda(valor_parcial_fases)}**."))

    # --- Preenche o contador de modalidade (posicionado acima, sob a Prova Real) ---
    # Todas as horas de campo são presenciais; somamos as etapas marcadas como presenciais.
    # O restante das horas (etapas não presenciais do Plano Detalhado) é online.
    horas_presenciais = total_horas_campo + horas_fases_presencial
    horas_online = horas_fases_online
    with placeholder_modalidade:
        st.markdown("#### ⏱️ Modalidade das Horas")
        cm1, cm2 = st.columns(2)
        cm1.metric("🟢 Horas Presenciais", f"{horas_presenciais} h")
        cm2.metric("💻 Horas Online", f"{horas_online} h")
        st.caption(
            "Presenciais = todas as horas de campo (H&M, entrevistas, grupos focais, OAC, visitas "
            "técnicas e aprofundamento) + etapas do Plano Detalhado marcadas como presenciais. "
            "Online = etapas do Plano Detalhado não marcadas como presenciais."
        )

    st.markdown("---")
    st.markdown("### 💰 4. Precificação Final e Logística")

    horas_totais = total_horas_campo + total_horas_fases
    valor_op1 = horas_totais * taxa_hora

    st.write("#### Planejamento de Logística (Opção 2)")
    st.session_state.logistica_dados = render_logistica(
        valor_op1, key_prefix="dcs",
        percentual_padrao=config.PERCENTUAL_LOG_DCS, dias_padrao=config.DIAS_IDA_DCS
    )
    logistica_total = st.session_state.logistica_dados["total"]
    valor_op2 = valor_op1 + logistica_total

    # Salvamento na memória global
    st.session_state.valores_finais["op1"] = valor_op1
    st.session_state.valores_finais["op2"] = valor_op2
    st.session_state.valores_finais["horas_totais"] = horas_totais
    st.session_state.valores_finais["taxa_hora"] = taxa_hora
    st.session_state.valores_finais["horas_presenciais"] = horas_presenciais
    st.session_state.valores_finais["horas_online"] = horas_online

    st.write(f"**Racional da Precificação Técnica:** {total_horas_campo}h (Campo) + {total_horas_fases}h (Plano Detalhado) = **{horas_totais} horas totais de projeto**.")

    c_tot1, c_tot2 = st.columns(2)
    c_tot1.metric("Total OP1 (Serviço Técnico)", formatar_moeda(valor_op1))
    c_tot2.metric("Total OP2 (Com Logística)", formatar_moeda(valor_op2))

    st.write("")
    st.button("Salvar e Avançar para Proposta Técnica ➡️", on_click=ir_para,
              args=("📝 2. Proposta Técnica",), type="primary")


def render_tecnica():
    st.title("📝 2. Gerador de Proposta Técnica")
    st.caption("As informações inseridas aqui alimentarão automaticamente a Proposta Comercial.")

    with st.sidebar:
        st.markdown("---")
        template_upload = st.file_uploader("Suba o template (Técnica)", type="pptx")

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

        c7, c8, c9 = st.columns([0.3, 0.4, 0.3])
        st.session_state.memoria_geral["formato"] = c7.selectbox("Formato ({{FORMATO}})*", ["Híbrido", "Presencial", "Online"], index=["Híbrido", "Presencial", "Online"].index(st.session_state.memoria_geral.get("formato", "Híbrido")))

        idiomas_selecionados = c8.multiselect("Idioma ({{IDIOMA}})*", ["Português", "Espanhol", "Inglês"], default=["Português"])
        if len(idiomas_selecionados) == 1:
            idioma_str = idiomas_selecionados[0]
        elif len(idiomas_selecionados) == 2:
            idioma_str = f"{idiomas_selecionados[0]} e {idiomas_selecionados[1]}"
        elif len(idiomas_selecionados) > 2:
            idioma_str = ", ".join(idiomas_selecionados[:-1]) + f" e {idiomas_selecionados[-1]}"
        else:
            idioma_str = ""
        st.session_state.memoria_geral["idioma"] = idioma_str

        # idas: unificado com a logística. Pré-preenche com o que veio da Precificação,
        # mas o consultor pode sobrescrever.
        idas_sugerido = st.session_state.get("logistica_dados", {}).get("idas", 0)
        idas_atual = st.session_state.memoria_geral.get("idas", 0)
        valor_idas_inicial = idas_atual if idas_atual else idas_sugerido
        st.session_state.memoria_geral["idas"] = c9.number_input(
            "Nº de Idas Presenciais ({{IDAS}})", min_value=0, value=int(valor_idas_inicial)
        )
        if idas_sugerido and idas_sugerido != st.session_state.memoria_geral["idas"]:
            c9.caption(f"Logística sugeriu {idas_sugerido} ida(s).")

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
                "{{IDIOMA}}": st.session_state.memoria_geral["idioma"],
                "{{IDAS}}": str(st.session_state.memoria_geral["idas"]),
                "{{N_LID}}": str(n_lid_total),
                "{{N_OPER}}": str(n_oper),
                "{{N_PTERC}}": str(n_p_terc)
            }
            with st.spinner("Construindo arquivo..."):
                st.session_state.pptx_gerado, avisos = processar_apresentacao(
                    template_upload, mapa, atividades_lista, "Técnica", None, qtd_meses_projeto
                )
                st.session_state.nome_arquivo = f"Tecnica_{st.session_state.memoria_geral['cliente']}.pptx"
            for aviso in avisos:
                st.warning(aviso)
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

        st.markdown(esc_md(f"**Total Investimento Técnico (OP1):** {formatar_moeda(v_op1)}"))
        st.markdown(esc_md(f"**Total Turnkey com Logística (OP2):** {formatar_moeda(v_op2)}"))

        st.write("---")
        qtd_parcelas = st.number_input("Quantidade de Parcelas ({{QTD_PARCELAS}})", min_value=1, value=12)
        st.session_state.valores_finais["qtd_parcelas"] = qtd_parcelas

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
                "{{IDIOMA}}": st.session_state.memoria_geral["idioma"],
                "{{IDAS}}": str(st.session_state.memoria_geral["idas"]),
                "{{VALOR_OP1}}": formatar_moeda(v_op1),
                "{{VALOR_OP2}}": formatar_moeda(v_op2),
                "{{VALOR_OP1_EXT}}": valor_por_extenso(v_op1),
                "{{VALOR_OP2_EXT}}": valor_por_extenso(v_op2),
                "{{QTD_PARCELAS}}": str(qtd_parcelas),
                "{{VLR1_PARCELAS}}": formatar_moeda(v_op1 / qtd_parcelas) if qtd_parcelas > 0 else "0",
                "{{VLR2_PARCELAS}}": formatar_moeda(v_op2 / qtd_parcelas) if qtd_parcelas > 0 else "0"
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
                st.session_state.pptx_gerado, avisos = processar_apresentacao(
                    template_upload, mapa_comercial, [], "Comercial", dados_financeiros, 12
                )
                st.session_state.nome_arquivo = f"Comercial_{st.session_state.memoria_geral['cliente']}.pptx"
            for aviso in avisos:
                st.warning(aviso)
            st.success("Comercial gerada com sucesso!")
            st.session_state.tentou_gerar = False

    if st.session_state.pptx_gerado and st.session_state.nome_arquivo.startswith("Comercial"):
        st.download_button("⬇️ Baixar PPTX Comercial", data=st.session_state.pptx_gerado, file_name=st.session_state.nome_arquivo)
        st.write("")
        st.button("Avançar para Handover (Operações) ➡️", on_click=ir_para, args=("🤝 4. Handover (Operações)",), type="primary")
