import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
import io
import datetime

# --- 1. FUNÇÕES TÉCNICAS E GANTT ---

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
        return 1072 if "Mapeamento" in tipo_rps else 1606
    elif servico == "Pulse":
        return 56
    elif servico == "Pontuais / Palestras":
        tabela_p = {"Palestra Online": 30, "Palestra Presencial": 36, "Imersão Liderança": 40}
        return tabela_p.get(tipo_pontual, 16)
    return 0

def processar_pptx(template_file, mapeamento, atividades):
    prs = Presentation(template_file)
    
    # 1. Substituição de Texto Simples
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                for key, value in mapeamento.items():
                    if key in shape.text:
                        shape.text = shape.text.replace(key, str(value))
            
            # 2. Lógica do Gantt (Procura a tabela 12x12)
            if shape.has_table:
                tbl = shape.table
                # Verifica se é a nossa tabela de avanço (12 colunas de meses + 1 de atividade)
                if len(tbl.columns) >= 12:
                    for row_idx, atividade in enumerate(atividades):
                        if row_idx + 1 < len(tbl.rows): # Evita estourar a tabela
                            row = tbl.rows[row_idx + 1]
                            row.cells[0].text = atividade['nome']
                            for mes in atividade['meses']:
                                if mes < len(tbl.columns):
                                    cell = row.cells[mes]
                                    cell.fill.solid()
                                    cell.fill.foreground_color.rgb = RGBColor(0, 128, 0) # Verde CPTO
    
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# --- 2. INTERFACE STREAMLIT ---
st.set_page_config(page_title="Emissor CPTO v6.0", layout="wide")
st.title("🎯 Emissor de Propostas com Upload Direto")

# SIDEBAR COM UPLOAD (O BYPASS!)
with st.sidebar:
    st.header("📁 Template Original")
    template_upload = st.file_uploader("Suba o arquivo .pptx com as {{TAGS}}", type="pptx")
    st.markdown("---")
    st.header("💰 Financeiro")
    valor_hora = st.number_input("Valor Hora (R$)", value=480.0)
    entidade = st.selectbox("Imposto:", ["Comportamento (20%)", "Escola (11%)"])
    imposto = 0.20 if "20%" in entidade else 0.11

# --- 3. FORMULÁRIO ---
tab1, tab2, tab3 = st.tabs(["📄 Identificação", "👥 Público", "📅 Gantt"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        servico = st.selectbox("Serviço Principal", ["Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)", "Riscos Psicossociais (RPS)", "Pulse", "Pontuais / Palestras"])
        cliente = st.text_input("Nome da Empresa")
        unidade = st.text_input("Unidade")
        num_prop = st.text_input("Nº Proposta")
    with col2:
        formato = st.selectbox("Formato", ["Híbrido", "Presencial", "Online"])
        idioma = st.selectbox("Idioma", ["Português", "Espanhol", "Inglês"])
        prazo = st.text_input("Prazo", "8 meses")

    justificativa = st.text_area("Justificativa")
    objetivo = st.text_area("Objetivo")

with tab2:
    c1, c2, c3 = st.columns(3)
    n_pr = c1.number_input("N° Executivos", value=0)
    n_exec = c2.number_input("N° Alta Liderança", value=0)
    n_coord = c3.number_input("N° Coordenadores", value=0)
    n_super = c1.number_input("N° Supervisores", value=0)
    n_sec = c2.number_input("N° Segurança", value=0)
    n_oper = c3.number_input("N° Operacionais", value=0)
    n_col3 = c1.number_input("N° Terceiros", value=0)
    n_lid3 = c2.number_input("N° Líderes Terceiros", value=0)

    n_lid_total = n_pr + n_exec + n_coord + n_super
    n_prop = n_lid_total + n_sec + n_oper
    n_p_terc = n_prop + n_col3 + n_lid3

with tab3:
    st.write("Preencha as fases e selecione os meses:")
    atividades_lista = []
    for i in range(8):
        ca, cm = st.columns([0.4, 0.6])
        nome_at = ca.text_input(f"Fase {i+1}", key=f"f{i}")
        meses_at = cm.multiselect(f"Meses Ativa", list(range(1, 13)), key=f"m{i}")
        if nome_at: atividades_lista.append({"nome": nome_at, "meses": meses_at})

# --- 4. CÁLCULOS E GERAÇÃO ---
tipo_rps = ""
tipo_pontual = ""
if servico == "Riscos Psicossociais (RPS)": tipo_rps = st.radio("Escopo", ["Mapeamento", "Gestão"])
if servico == "Pontuais / Palestras": tipo_pontual = st.selectbox("Tipo", ["Palestra Online", "Palestra Presencial", "Imersão Liderança"])

ch_calculada = calcular_esforço(servico, n_p_terc, n_lid_total, tipo_rps, tipo_pontual)
investimento = (ch_calculada * valor_hora) / (1 - imposto)

st.markdown("---")
if st.button("🔥 CALCULAR E GERAR PROPOSTA"):
    if not template_upload:
        st.warning("⚠️ Por favor, suba o arquivo de template na barra lateral primeiro!")
    else:
        mapeamento = {
            "{{CLIENTE}}": cliente, "{{UNIDADE}}": unidade, "{{NUM_PROP}}": num_prop,
            "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
            "{{JUSTIFICATIVA}}": justificativa, "{{OBJETIVO}}": objetivo,
            "{{PUBLICO}}": n_p_terc, "{{PRAZO}}": prazo, "{{FORMATO}}": formato, "{{IDIOMA}}": idioma,
            "{{N_PR}}": n_pr, "{{N_EXEC}}": n_exec, "{{N_COORD}}": n_coord, "{{N_SUPER}}": n_super,
            "{{N_LID}}": n_lid_total, "{{N_SEC}}": n_sec, "{{N_OPER}}": n_oper, "{{N_PROP}}": n_prop,
            "{{N_COL3}}": n_col3, "{{N_LID3}}": n_lid3, "{{N_PTERC}}": n_p_terc
        }
        
        arquivo_gerado = processar_pptx(template_upload, mapeamento, atividades_lista)
        
        st.success("✅ Proposta processada com sucesso!")
        st.download_button(
            label="⬇️ Baixar PowerPoint Pronto",
            data=arquivo_gerado,
            file_name=f"Proposta_{cliente}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
