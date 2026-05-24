import streamlit as st
import pandas as pd
from ferramentas.utilidades import formatar_moeda

def render_handover():
    st.title("🤝 4. Handover de Projeto (Operações)")
    st.info("Finalize as informações abaixo para consolidar o Recibo do Projeto para a equipe de Operações.")
    
    memoria = st.session_state.memoria_geral
    finais = st.session_state.valores_finais
    logistica = st.session_state.logistica_dados
    
    # Busca nomes das unidades dependendo do serviço
    if st.session_state.servico_selecionado == "Proposta Pontual (Palestras/Workshops)":
        unidades_str = ", ".join([u["nome"] for u in st.session_state.unidades_pontual])
    else:
        unidades_str = ", ".join([u["nome"] for u in st.session_state.unidades_dcs])

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
    
    # Regra do valor por deslocamento
    valor_deslocamento = "R$ 0,00"
    if logistica["idas"] > 0:
        valor_deslocamento = formatar_moeda(logistica["total"] / logistica["idas"])
    
    # Construção da visualização da Logística conforme regra
    texto_logistica = f"Tipo: {logistica['tipo']}\n"
    if "Completa" in logistica['tipo']:
        for item, valor in logistica['detalhes'].items():
            texto_logistica += f" - {item}: {formatar_moeda(valor)}\n"
        texto_logistica += f"\nCusto Total Logística: {formatar_moeda(logistica['total'])}"
    elif "Estimada" in logistica['tipo']:
        texto_logistica += f"Margem aplicada: {logistica['detalhes'].get('Percentual Aplicado', '')} \nCusto Total Estimado: {formatar_moeda(logistica['total'])}"
    elif "Base" in logistica['tipo']:
        texto_logistica += f" - Táxi/Alimentação \nCusto Total Base: {formatar_moeda(logistica['total'])}"
    else:
        texto_logistica += "Cliente assume todos os custos."

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
            memoria["num_prop"], memoria["cliente"], unidades_str, memoria["servico"], memoria["idioma"],
            f"{finais['horas_totais']} horas", memoria["formato"], formatar_moeda(finais["op2"]), formatar_moeda(finais["taxa_hora"]),
            cc_alim, cc_hotel, f"{logistica['idas']} deslocamento(s)",
            memoria["prazo"], prazo_inicio, caminho_drive, lanc_meta,
            f"{finais['qtd_parcelas']} Parcela(s)", observacoes, sugestao_reuniao, continuidade_gestor,
            contato_nome, contato_email, contato_tel
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
