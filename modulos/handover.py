import streamlit as st
import pandas as pd
from ferramentas.utilidades import formatar_moeda

def render_handover():
    st.title("🤝 4. Handover de Projeto (Operações)")
    st.info("Finalize as informações abaixo para consolidar o Recibo do Projeto para a equipe de Operações.")
    
    # O .get() é a nossa rede de segurança. Se não achar, puxa um vazio/zero.
    memoria = st.session_state.get("memoria_geral", {})
    finais = st.session_state.get("valores_finais", {})
    logistica = st.session_state.get("logistica_dados", {"tipo": "Não informado", "total": 0.0, "idas": 0, "detalhes": {}})
    
    servico_atual = st.session_state.get("servico_selecionado", "Não selecionado")

    # Busca nomes das unidades dependendo do serviço com segurança
    if servico_atual == "Proposta Pontual (Palestras/Workshops)":
        unidades = st.session_state.get("unidades_pontual", [{"nome": "Não informada"}])
    else:
        unidades = st.session_state.get("unidades_dcs", [{"nome": "Não informada"}])
        
    unidades_str = ", ".join([u["nome"] for u in unidades])

    # --- CAMPOS MANUAIS DO COMERCIAL ---
    with st.expander("📝 1. Preenchimento de Dados Operacionais", expanded=True):
        c1, c2 = st.columns(2)
        cc_alim = c1.text_input("Centro de Custo (Alimentação, Táxi Base Consultor)", value="DP")
        cc_hotel = c2.text_input("Centro de Custo (Hotel, Aéreo e Táxi no Cliente)", value="DP")
        
        c3, c4 = st.columns(2)
        prazo_inicio = c3.text_input("Prazo negociado de início das atividades", placeholder="Ex: Não negociado")
        lanc_meta = c4.text_input("Lançar na meta", value="Comportamento")
        
        caminho_drive = st.text_input("Caminho da proposta no drive", placeholder="Comportamento Consultoria em Recursos Humanos LTDA\\COMERCIAL - Documentos\\...")
        observacoes = st.text_area("Observações (Ex: Pedido N: XXXXX)")
        
        c5, c6 = st.columns(2)
        sugestao_reuniao = c5.selectbox("Sugestão do comercial para reunião", ["Alinhamento com o comercial", "Reunião de abertura", "Alinhamento administrativo", "Reunião fracionada"])
        continuidade_gestor = c6.text_input("Continuidade de cliente. Gestor", placeholder="Rafael / Jivago")
        
        st.markdown("##### Dados do Cliente")
        c7, c8, c9 = st.columns(3)
        contato_nome = c7.text_input("Contato (Nome)")
        contato_email = c8.text_input("E-mail")
        contato_tel = c9.text_input("Telefone do contato")

    # --- GERAÇÃO DA TABELA VISUAL ---
    st.markdown("### 📋 Resumo Consolidado (Recibo)")
    
    idas = logistica.get("idas", 0)
    total_log = logistica.get("total", 0.0)
    
    # Regra do valor por deslocamento
    valor_deslocamento = "R$ 0,00"
    if idas > 0:
        valor_deslocamento = formatar_moeda(total_log / idas)
    
    # Construção da visualização da Logística conforme regra
    tipo_log = logistica.get("tipo", "")
    texto_logistica = f"Tipo: {tipo_log}\n"
    
    if "Completa" in tipo_log:
        detalhes = logistica.get('detalhes', {})
        for item, valor in detalhes.items():
            texto_logistica += f" - {item}: {formatar_moeda(valor)}\n"
        texto_logistica += f"\nCusto Total Logística: {formatar_moeda(total_log)}"
    elif "Estimada" in tipo_log:
        perc = logistica.get('detalhes', {}).get('Percentual Aplicado', '')
        texto_logistica += f"Margem aplicada: {perc} \nCusto Total Estimado: {formatar_moeda(total_log)}"
    elif "Base" in tipo_log:
        texto_logistica += f" - Táxi/Alimentação \nCusto Total Base: {formatar_moeda(total_log)}"
    else:
        texto_logistica += "Cliente assume todos os custos (ou Sem Logística)."

    # Tabela principal de Informações Comerciais
    dados_tabela = {
        "Campo": [
            "Nº da proposta", "Cliente", "Unidades de atendimento", "Serviço", "Idioma",
            "Carga Horária Total (CHT)", "Formato", "Valor total do projeto", "Valor hora",
            "Centro de Custo (Base)", "Centro de Custo (Cliente)", "Quantidade de idas presenciais",
            "Prazo de execução", "Prazo de início", "Caminho no drive", "Lançar na meta",
            "Pagamento", "Observações", "Sugestão p/ reunião", "Gestor de Continuidade",
            "Contato", "E-mail", "Telefone"
        ],
        "Informação": [
            memoria.get("num_prop", ""), 
            memoria.get("cliente", ""), 
            unidades_str, 
            servico_atual, 
            memoria.get("idioma", "Português"),
            f"{finais.get('horas_totais', 0)} horas", 
            memoria.get("formato", "Híbrido"), 
            formatar_moeda(finais.get("op2", 0.0)), 
            formatar_moeda(finais.get("taxa_hora", 0.0)),
            cc_alim, cc_hotel, 
            f"{idas} deslocamento(s)",
            memoria.get("prazo", ""), 
            prazo_inicio, 
            caminho_drive, 
            lanc_meta,
            f"{finais.get('qtd_parcelas', 1)} Parcela(s)", 
            observacoes, 
            sugestao_reuniao, 
            continuidade_gestor,
            contato_nome, 
            contato_email, 
            contato_tel
        ]
    }
    
    st.table(pd.DataFrame(dados_tabela))
    
    # Tabela exclusiva de Logística
    st.markdown("#### Detalhamento Logístico")
    dados_logistica = {
        "Resumo do Cálculo Logístico": [texto_logistica],
        "Valor por Deslocamento (Total / Idas)": [valor_deslocamento]
    }
    st.table(pd.DataFrame(dados_logistica))

    # Botão apenas visual para a aprovação futura
    st.success("✅ Recibo pronto! Selecione o texto acima e copie para a sua planilha oficial ou salve a página em PDF.")
