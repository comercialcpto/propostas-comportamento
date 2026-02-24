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

# --- FUNÇÃO AUXILIAR DE FORMATAÇÃO ---
def formatar_run(run, eh_capa, key):
    if eh_capa:
        run.font.name = "Annantason Expanded Bold"
        run.font.size = Pt(21)
        run.font.color.rgb = BRANCO
    elif key in ["{{PUBLICO}}", "{{IDIOMA}}"]:
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.color.rgb = BRANCO
    else:
        run.font.name = "Calibri"
        run.font.size = Pt(14)
        run.font.color.rgb = CINZA_ESCURO

# --- MOTOR DE SUBSTITUIÇÃO ROBUSTO ---
def substituir_texto_em_shape(shape, mapa, eh_capa):
    if hasattr(shape, "text_frame") and shape.text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                for key, value in mapa.items():
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
                        formatar_run(run, eh_capa, key)

def processar_pptx(template_file, mapa, atividades):
    prs = Presentation(template_file)
    
    for i, slide in enumerate(prs.slides):
        eh_capa = (i == 0)
        
        for shape in slide.shapes:
            # 1. SUBSTITUIÇÃO EM TEXTOS NORMAIS
            substituir_texto_em_shape(shape, mapa, eh_capa)
            
            # 2. SUBSTITUIÇÃO DENTRO DE TABELAS (Pág 7, 20, etc.)
            if shape.has_table:
                tbl = shape.table
                
                # Loop para substituir variáveis dentro das células da tabela
                for row in tbl.rows:
                    for cell in row.cells:
                        substituir_texto_em_shape(cell, mapa, eh_capa)

                # 3. LÓGICA DO GANTT (Específica para a tabela 12 colunas)
                # O cabeçalho é a linha 0, então começamos da linha 1
                if len(tbl.columns) >= 12:
                    # Pintar as células de verde conforme os meses selecionados
                    for row_idx, atividade in enumerate(atividades):
                        target_row = row_idx + 1 # Pula o header (mes 1, mes 2...)
                        if target_row < len(tbl.rows):
                            row = tbl.rows[target_row]
                            row.cells[0].text = atividade['nome']
                            # Reset de cor e aplicação do verde
                            for m_idx in range(1, len(tbl.columns)):
                                if m_idx in atividade['meses']:
                                    cell = row.cells[m_idx]
                                    cell.fill.solid()
                                    cell.fill.foreground_color.rgb = VERDE_CPTO

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Emissor CPTO v8.0", layout="wide")
st.title("🎯 Emissor de Propostas Grupo Comportamento - v8.0")

with st.sidebar:
    st.header("📁 Template")
    template_upload = st.file_uploader("Suba o template .pptx", type="pptx")
    st.markdown("---")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Faturar por:", ["Comportamento (20%)", "Escola (11%)"])
    imposto = 0.20 if "20%" in entidade else 0.11

tab1, tab2, tab3, tab4 = st.tabs(["📄 Identificação", "👥 Público", "📅 Cronograma", "📊 Financeiro"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        servico = st.selectbox("Serviço", ["Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)", "Riscos Psicossociais (RPS)", "Pulse", "Pontuais / Palestras"])
        cliente = st.text_input("Empresa", value="Teste Cliente")
        unidade = st.text_input("Unidade", value="São Paulo")
        num_prop = st.text_input("Nº Proposta", value="001/2026")
    with col2:
        formato = st.selectbox("Formato", ["Híbrido", "Presencial", "Online"])
        idioma = st.selectbox("Idioma", ["Português", "Espanhol", "Inglês"])
        prazo = st.text_input("Prazo", value="8 meses")
        idas_presenciais = st.number_input("Nº de Idas Presenciais", min_value=0, value=2)
    
    justificativa = st.text_area("Justificativa", value="Justificativa de teste para a proposta.")
    objetivo = st.text_area("Objetivo", value="Objetivo principal do projeto de consultoria.")

with tab2:
    st.write("Preencha a população para a Tabela da Página 7:")
    c1, c2, c3 = st.columns(3)
    n_pr = c1.number_input("N° Executivos", value=2)
    n_exec = c2.number_input("N° Alta Liderança", value=5)
    n_coord = c3.number_input("N° Coordenadores", value=10)
    n_super = c1.number_input("N° Supervisores", value=15)
    n_lid_total = n_pr + n_exec + n_coord + n_super
    n_sec = c2.number_input("N° Segurança", value=3)
    n_oper = c3.number_input("N° Operacionais", value=200)
    n_col3 = c1.number_input("N° Terceiros", value=50)
    n_lid3 = c2.number_input("N° Líderes Terceiros", value=5)
    
    n_prop = n_lid_total + n_sec + n_oper
    n_p_terc = n_prop + n_col3 + n_lid3

with tab3:
    st.write("Fases do Cronograma (Página 6):")
    atividades_lista = []
    for i in range(6):
        ca, cm = st.columns([0.4, 0.6])
        nome_at = ca.text_input(f"Atividade {i+1}", key=f"f_{i}")
        meses_at = cm.multiselect(f"Meses", list(range(1, 13)), key=f"m_{i}")
        if nome_at: atividades_lista.append({"nome": nome_at, "meses": meses_at})

with tab4:
    st.write("Relatórios e PTCs (Página 20):")
    cr1, cr2 = st.columns(2)
    qtd_rel = cr1.number_input("Qtd Unidades com Relatório", value=2)
    tem_corp = cr2.checkbox("Gerar Relatório Corporativo?", value=True)
    tot_rel = qtd_rel + (1 if tem_corp else 0)
    tot_plan = st.number_input("Total de Planos (PTC)", value=1)
    
    ch_total = 144 # Exemplo base Lote 4
    investimento = (ch_total * valor_hora) / (1 - imposto)
    st.metric("Investimento Total Estimado", f"R$ {investimento:,.2f}")

if st.button("🚀 GERAR PROPOSTA FINAL"):
    if not template_upload:
        st.error("Suba o template na barra lateral!")
    else:
        mapa_final = {
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
        
        pptx_io = processar_pptx(template_upload, mapa_final, atividades_lista)
        st.success("✅ PowerPoint gerado com sucesso!")
        st.download_button("⬇️ Baixar Arquivo", pptx_io, f"Proposta_{cliente}.pptx")
