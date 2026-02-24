import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
import io
import datetime

# --- CONFIGURAÇÕES DE CORES ---
BRANCO = RGBColor(255, 255, 255)
CINZA_ESCURO = RGBColor(64, 64, 64)
VERDE_CPTO = RGBColor(0, 128, 0)

# --- FUNÇÕES TÉCNICAS ---

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
        return 1072 if tipo_rps == "Mapeamento" else 1606
    elif servico == "Pulse":
        return 56
    elif servico == "Pontuais / Palestras":
        tabela_p = {"Palestra Online": 30, "Palestra Presencial": 36, "Imersão Liderança": 40}
        return tabela_p.get(tipo_pontual, 16)
    return 0

def aplicar_estilo(run, fonte_nome, tamanho, cor):
    run.font.name = fonte_nome
    run.font.size = Pt(tamanho)
    run.font.color.rgb = cor

def processar_pptx(template_file, mapeamento, atividades):
    prs = Presentation(template_file)
    
    for i, slide in enumerate(prs.slides):
        # slide index 0 costuma ser a capa
        eh_capa = (i == 0)
        
        for shape in slide.shapes:
            # 1. TRATAMENTO DE TEXTOS
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        texto_original = run.text
                        novo_texto = texto_original
                        
                        # Substitui as variáveis do mapeamento
                        for key, value in mapeamento.items():
                            if key in novo_texto:
                                novo_texto = novo_texto.replace(key, str(value))
                                run.text = novo_texto
                                
                                # APLICAÇÃO DE REGRAS DE DESIGN DO FEEDBACK
                                if eh_capa:
                                    aplicar_estilo(run, "Annantason Expanded Bold", 21, BRANCO)
                                elif key in ["{{PUBLICO}}", "{{IDIOMA}}"]:
                                    aplicar_estilo(run, "Calibri", 14, BRANCO)
                                else:
                                    # Resto do Charter e slides
                                    aplicar_estilo(run, "Calibri", 14, CINZA_ESCURO)

            # 2. LÓGICA DO GANTT (Tabelas)
            if shape.has_table:
                tbl = shape.table
                # Se a tabela tem 12 ou 13 colunas (Meses + Atividade)
                if len(tbl.columns) >= 12:
                    for row_idx, atividade in enumerate(atividades):
                        if row_idx + 1 < len(tbl.rows):
                            row = tbl.rows[row_idx + 1]
                            row.cells[0].text = atividade['nome']
                            # Limpa formatação anterior e pinta de verde os meses selecionados
                            for m in range(1, len(tbl.columns)):
                                if m in atividade['meses']:
                                    cell = row.cells[m]
                                    cell.fill.solid()
                                    cell.fill.foreground_color.rgb = VERDE_CPTO

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Emissor CPTO v7.0", layout="wide")
st.title("🎯 Emissor de Propostas Grupo Comportamento - v7.0")

with st.sidebar:
    st.header("📁 Configurações")
    template_upload = st.file_uploader("Suba o template .pptx", type="pptx")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Faturar por:", ["Comportamento (20%)", "Escola (11%)"])
    imposto = 0.20 if "20%" in entidade else 0.11

tab1, tab2, tab3, tab4 = st.tabs(["📄 Identificação", "👥 Público", "📅 Cronograma", "📊 Financeiro"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        servico = st.selectbox("Serviço", ["Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)", "Riscos Psicossociais (RPS)", "Pulse", "Pontuais / Palestras"])
        cliente = st.text_input("Empresa")
        unidade = st.text_input("Unidade")
        num_prop = st.text_input("Nº Proposta")
    with col2:
        formato = st.selectbox("Formato", ["Híbrido", "Presencial", "Online"])
        idioma = st.selectbox("Idioma", ["Português", "Espanhol", "Inglês"])
        prazo = st.text_input("Prazo", "8 meses")
        idas_presenciais = st.number_input("Nº de Idas Presenciais ({{IDAS}})", min_value=0, value=1)
    
    justificativa = st.text_area("Justificativa")
    objetivo = st.text_area("Objetivo")

with tab2:
    c1, c2, c3 = st.columns(3)
    n_pr = c1.number_input("N° Executivos (Pres/VP/Dir)", value=0)
    n_exec = c2.number_input("N° Alta Liderança (Gerentes)", value=0)
    n_coord = c3.number_input("N° Coordenadores", value=0)
    n_super = c1.number_input("N° Supervisores", value=0)
    n_sec = c2.number_input("N° Segurança", value=0)
    n_oper = c3.number_input("N° Operacionais", value=0)
    n_col3 = c1.number_input("N° Colab. Terceiros", value=0)
    n_lid3 = c2.number_input("N° Líderes Terceiros", value=0)
    
    n_lid_total = n_pr + n_exec + n_coord + n_super
    n_prop = n_lid_total + n_sec + n_oper
    n_p_terc = n_prop + n_col3 + n_lid3

with tab3:
    st.write("Fases do Avanço Mensal:")
    atividades_lista = []
    for i in range(8):
        ca, cm = st.columns([0.4, 0.6])
        nome_at = ca.text_input(f"Fase {i+1}", key=f"fase_{i}")
        meses_at = cm.multiselect(f"Meses", list(range(1, 13)), key=f"mes_{i}")
        if nome_at: atividades_lista.append({"nome": nome_at, "meses": meses_at})

with tab4:
    st.subheader("Relatórios e PTCs")
    cr1, cr2 = st.columns(2)
    qtd_rel = cr1.number_input("Qtd Unidades com Relatório", value=1)
    tem_corp = cr2.checkbox("Relatório Corporativo?")
    tot_rel = qtd_rel + (1 if tem_corp else 0)
    tot_plan = st.number_input("Total de Planos (PTC)", value=1)
    
    tipo_rps = st.radio("Escopo RPS", ["Mapeamento", "Gestão"]) if "RPS" in servico else ""
    tipo_pontual = st.selectbox("Tipo Evento", ["Palestra Online", "Palestra Presencial", "Imersão Liderança"]) if "Pontuais" in servico else ""
    
    ch_total = calcular_esforço(servico, n_p_terc, n_lid_total, tipo_rps, tipo_pontual)
    investimento = (ch_total * valor_hora) / (1 - imposto)
    st.metric("Investimento Total", f"R$ {investimento:,.2f}")

if st.button("🚀 GERAR PROPOSTA FINAL"):
    if not template_upload:
        st.error("Suba o template na barra lateral!")
    else:
        # MAPA DE TRADUÇÃO COMPLETO (Garantindo que todos os campos do feedback existam)
        mapa = {
            "{{CLIENTE}}": cliente, "{{UNIDADE}}": unidade, "{{NUM_PROP}}": num_prop,
            "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
            "{{JUSTIFICATIVA}}": justificativa, "{{OBJETIVO}}": objetivo,
            "{{PUBLICO}}": str(n_p_terc), "{{PRAZO}}": prazo, "{{FORMATO}}": formato, "{{IDIOMA}}": idioma,
            "{{N_PR}}": str(n_pr), "{{N_EXEC}}": str(n_exec), "{{N_COORD}}": str(n_coord), "{{N_SUPER}}": str(n_super),
            "{{N_LID}}": str(n_lid_total), "{{N_SEC}}": str(n_sec), "{{N_OPER}}": str(n_oper),
            "{{N_PROP}}": str(n_prop), "{{N_COL3}}": str(n_col3), "{{N_LID3}}": str(n_lid3),
            "{{N_PTERC}}": str(n_p_terc), "{{IDAS}}": str(idas_presenciais),
            "{{TOT_REL}}": str(tot_rel), "{{QTD_REL}}": str(qtd_rel), "{{TOT_PLAN}}": str(tot_plan)
        }
        
        pptx_pronto = processar_pptx(template_upload, mapa, atividades_lista)
        st.success("✅ Processado!")
        st.download_button("⬇️ Baixar Proposta", pptx_pronto, f"Proposta_{cliente}.pptx")
