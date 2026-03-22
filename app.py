import streamlit as st
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
import io
import datetime

# ==========================================
# 1. CONFIGURAÇÕES E FUNÇÕES AUXILIARES
# ==========================================
BRANCO = RGBColor(255, 255, 255)
CINZA_ESCURO = RGBColor(64, 64, 64)
VERDE_CPTO = RGBColor(0, 128, 0)

def formatar_run(run, eh_capa, key):
    if eh_capa or key == "{{ESCOPO}}":
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

def valor_por_extenso(valor):
    # Função simplificada para escrever o valor em reais (suporta até 999 mil para propostas padrão)
    # Se o número passar de 1 milhão, ele fará uma aproximação elegante.
    unidades = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
    dez_a_dezenove = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    dezenas = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta", "oitenta", "noventa"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]

    def converter_trio(n):
        if n == 100: return "cem"
        if n == 0: return ""
        c, resto = n // 100, n % 100
        d, u = resto // 10, resto % 10
        partes = []
        if c > 0: partes.append(centenas[c])
        if d == 1: partes.append(dez_a_dezenove[u])
        else:
            if d > 1: partes.append(dezenas[d])
            if u > 0: partes.append(unidades[u])
        return " e ".join(partes)

    inteiro = int(valor)
    if inteiro == 0: return "zero reais"
    
    milhoes = inteiro // 1000000
    milhares = (inteiro % 1000000) // 1000
    resto = inteiro % 1000
    
    resultado = []
    if milhoes > 0:
        resultado.append(converter_trio(milhoes) + (" milhões" if milhoes > 1 else " milhão"))
    if milhares > 0:
        resultado.append(converter_trio(milhares) + " mil")
    if resto > 0:
        resultado.append(converter_trio(resto))
        
    extenso_final = ", ".join(resultado).replace(", e", " e")
    return extenso_final.capitalize() + (" reais" if inteiro > 1 else " real")

def calcular_amortizacao(qtd_parcelas):
    # Lógica de peso: 20, 20, 10, 10 e o resto 5. 
    pesos_base = [20, 20, 10, 10] + [5] * 30
    pesos_ativos = pesos_base[:qtd_parcelas]
    soma = sum(pesos_ativos)
    percentuais = [round((p / soma) * 100) for p in pesos_ativos]
    
    # Ajusta diferença de arredondamento no primeiro mês
    diferenca = 100 - sum(percentuais)
    if diferenca != 0:
        percentuais[0] += diferenca
    return percentuais

# ==========================================
# 2. MOTOR DE PROCESSAMENTO PPTX
# ==========================================
def deletar_slide(prs, slide):
    id_dict = { s.id: [i, s.rId] for i, s in enumerate(prs.slides._sldIdLst) }
    prs.part.drop_rel(id_dict[slide.slide_id][1])
    del prs.slides._sldIdLst[id_dict[slide.slide_id][0]]

def remover_linha_tabela(table, row_idx):
    tr = table.rows[row_idx]._tr
    tr.getparent().remove(tr)

def processar_apresentacao(template_file, mapa, atividades, tipo_doc, dados_fin=None):
    prs = Presentation(template_file)
    slides_para_deletar = []

    for i, slide in enumerate(prs.slides):
        eh_capa = (i == 0)
        deletar_este_slide = False
        
        for shape in slide.shapes:
            # 1. Faxineiro de Slides "Para DCS"
            if hasattr(shape, "text") and "Para DCS" in shape.text:
                if mapa.get("{{SERVICO}}", "") != "Diagnóstico (DCS/Clima/DCMA)":
                    deletar_este_slide = True
                    break # Não precisa olhar mais nada neste slide

            # 2. Substituição de Texto Normal
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in mapa.items():
                            if key in run.text:
                                run.text = run.text.replace(key, str(value))
                                formatar_run(run, eh_capa, key)
            
            # 3. Tratamento de Tabelas
            if shape.has_table:
                tbl = shape.table
                # Textos dentro da tabela
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for p in cell.text_frame.paragraphs:
                                for run in p.runs:
                                    for key, value in mapa.items():
                                        if key in run.text:
                                            run.text = run.text.replace(key, str(value))
                                            formatar_run(run, eh_capa, key)

                # Gantt (12 ou mais colunas)
                if len(tbl.columns) >= 12 and len(atividades) > 0:
                    for row_idx, atividade in enumerate(atividades):
                        target_row = row_idx + 1 
                        if target_row < len(tbl.rows):
                            row = tbl.rows[target_row]
                            row.cells[0].text = atividade['nome']
                            for m_idx in range(1, len(tbl.columns)):
                                if m_idx in atividade['meses']:
                                    cell = row.cells[m_idx]
                                    cell.fill.solid()
                                    cell.fill.fore_color.rgb = VERDE_CPTO

                # Tabela Financeira Comercial (Somente se dados_fin existir)
                if tipo_doc == "Comercial" and dados_fin:
                    try:
                        cabecalho = tbl.rows[0].cells[0].text.strip().lower()
                        
                        # Tabela 1: Macro Ações
                        if "macro" in cabecalho:
                            acoes = dados_fin['acoes']
                            linhas_para_deletar = []
                            for idx in range(1, len(tbl.rows)):
                                if idx <= len(acoes):
                                    tbl.rows[idx].cells[0].text = acoes[idx-1]['nome']
                                    tbl.rows[idx].cells[1].text = f"R$ {acoes[idx-1]['v1']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                    tbl.rows[idx].cells[2].text = f"R$ {acoes[idx-1]['v2']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                elif "Investimento total" not in tbl.rows[idx].cells[0].text:
                                    linhas_para_deletar.append(idx)
                            
                            for idx in reversed(linhas_para_deletar):
                                remover_linha_tabela(tbl, idx)

                        # Tabela 2: Faturamentos (Parcelas)
                        elif "meses" in cabecalho:
                            parcelas = dados_fin['parcelas']
                            val_mensal = dados_fin['val_parcela']
                            linhas_para_deletar = []
                            for idx in range(1, len(tbl.rows)):
                                if idx <= len(parcelas):
                                    tbl.rows[idx].cells[0].text = f"M{idx}"
                                    tbl.rows[idx].cells[1].text = f"{parcelas[idx-1]}%"
                                    val_calc = dados_fin['total_op2'] * (parcelas[idx-1] / 100)
                                    tbl.rows[idx].cells[2].text = f"R$ {val_calc:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                                elif "Total" not in tbl.rows[idx].cells[0].text:
                                    linhas_para_deletar.append(idx)
                            
                            for idx in reversed(linhas_para_deletar):
                                remover_linha_tabela(tbl, idx)
                    except:
                        pass # Proteção contra tabelas fora do padrão

        if deletar_este_slide:
            slides_para_deletar.append(slide)

    # Executa a limpeza dos slides DCS se necessário
    for slide in slides_para_deletar:
        deletar_slide(prs, slide)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. INTERFACE DE USUÁRIO E ROTEAMENTO
# ==========================================
st.set_page_config(page_title="Sistema Comportamento", layout="wide")

st.sidebar.title("🧭 Navegação")
menu = st.sidebar.radio("Selecione o Módulo:", ["🏠 Início", "💰 Criar Precificação", "📊 Criar Apresentação"])

if menu == "🏠 Início":
    st.title("Bem-vindo ao Sistema Comercial - Comportamento")
    st.write("Selecione um módulo no menu lateral para começar.")

elif menu == "💰 Criar Precificação":
    st.title("💰 Módulo de Precificação")
    st.info("🚧 Em construção. Aqui entrarão os cálculos baseados nas suas planilhas Excel.")

elif menu == "📊 Criar Apresentação":
    st.title("📊 Gerador de Propostas")
    tipo_apresentacao = st.radio("Selecione o tipo de documento:", ["Apresentação Técnica", "Apresentação Comercial"], horizontal=True)
    
    with st.sidebar:
        st.markdown("---")
        st.header("📁 Template Original")
        template_upload = st.file_uploader(f"Suba o template .pptx ({tipo_apresentacao})", type="pptx")
    
    # --- VARIÁVEIS COMUNS ---
    with st.expander("📍 1. Identificação Geral", expanded=True):
        c1, c2, c3 = st.columns(3)
        servico = c1.selectbox("Serviço Principal", ["Diagnóstico (DCS/Clima/DCMA)", "Mapeamento de Liderança (MPL)", "Riscos Psicossociais (RPS)", "Pulse", "Pontuais / Palestras"])
        cliente = c2.text_input("Nome da Empresa ({{CLIENTE}})")
        unidade = c3.text_input("Unidade ({{UNIDADE}})")
        
        c4, c5, c6 = st.columns(3)
        num_prop = c4.text_input("Nº da Proposta ({{NUM_PROP}})")
        escopo_tag = c5.text_input("Título do Escopo ({{ESCOPO}})", value="Encontro de Planejamento 2026")
        prazo = c6.text_input("Prazo ({{PRAZO}})")
        
        c7, c8, c9 = st.columns(3)
        formato = c7.selectbox("Formato ({{FORMATO}})", ["Híbrido", "Presencial", "Online"])
        idioma = c8.selectbox("Idioma ({{IDIOMA}})", ["Português", "Espanhol", "Inglês"])
        idas = c9.number_input("Nº de Idas Presenciais ({{IDAS}})", min_value=0)
        
        justificativa = st.text_area("Justificativa ({{JUSTIFICATIVA}})")
        objetivo = st.text_area("Objetivo ({{OBJETIVO}})")

    with st.expander("📅 2. Cronograma de Avanço (Gantt)"):
        atividades_lista = []
        for i in range(10):
            ca, cm = st.columns([0.4, 0.6])
            nome_at = ca.text_input(f"Fase {i+1}", key=f"f_{i}")
            meses_at = cm.multiselect("Meses", list(range(1, 13)), key=f"m_{i}")
            if nome_at: atividades_lista.append({"nome": nome_at, "meses": meses_at})

    # ==========================================
    # LÓGICA: APRESENTAÇÃO TÉCNICA
    # ==========================================
    if tipo_apresentacao == "Apresentação Técnica":
        with st.expander("👥 3. Detalhamento do Público e Relatórios"):
            cp1, cp2, cp3 = st.columns(3)
            n_pr = cp1.number_input("Executivos", value=0)
            n_exec = cp2.number_input("Alta Liderança", value=0)
            n_coord = cp3.number_input("Coordenadores", value=0)
            n_super = cp1.number_input("Supervisores", value=0)
            n_lid_extra = cp2.number_input("Outros Líderes", value=0)
            n_sec = cp3.number_input("Segurança", value=0)
            n_oper = cp1.number_input("Operacional", value=0)
            n_col3 = cp2.number_input("Terceiros", value=0)
            n_lid3 = cp3.number_input("Líderes Terc.", value=0)
            
            n_lid_total = n_pr + n_exec + n_coord + n_super + n_lid_extra
            n_prop = n_lid_total + n_sec + n_oper
            n_p_terc = n_prop + n_col3 + n_lid3

            st.write("---")
            cr1, cr2, cr3 = st.columns(3)
            qtd_rel = cr1.number_input("Qtd de Unidades com Relatório", value=1)
            tem_corp = cr2.checkbox("Relatório Corporativo?")
            tot_rel = qtd_rel + (1 if tem_corp else 0)
            tot_plan = cr3.number_input("Total de PTCs", value=1)

        if st.button("🚀 GERAR PROPOSTA TÉCNICA"):
            if template_upload:
                mapa = {
                    "{{SERVICO}}": servico, "{{CLIENTE}}": cliente, "{{UNIDADE}}": unidade, 
                    "{{NUM_PROP}}": num_prop, "{{ESCOPO}}": escopo_tag,
                    "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
                    "{{JUSTIFICATIVA}}": justificativa, "{{OBJETIVO}}": objetivo,
                    "{{PUBLICO}}": str(n_p_terc), "{{PRAZO}}": prazo, "{{FORMATO}}": formato, 
                    "{{IDIOMA}}": idioma, "{{IDAS}}": str(idas),
                    "{{N_PR}}": str(n_pr), "{{N_EXEC}}": str(n_exec), "{{N_COORD}}": str(n_coord), 
                    "{{N_SUPER}}": str(n_super), "{{N_LID}}": str(n_lid_total), "{{N_SEC}}": str(n_sec), 
                    "{{N_OPER}}": str(n_oper), "{{N_PROP}}": str(n_prop), "{{N_COL3}}": str(n_col3), 
                    "{{N_LID3}}": str(n_lid3), "{{N_PTERC}}": str(n_p_terc),
                    "{{TOT_REL}}": str(tot_rel), "{{QTD_REL}}": str(qtd_rel), "{{TOT_PLAN}}": str(tot_plan)
                }
                pptx_io = processar_apresentacao(template_upload, mapa, atividades_lista, "Técnica")
                st.success("Técnica gerada com sucesso!")
                st.download_button("⬇️ Baixar", pptx_io, f"Tecnica_{cliente}.pptx")
            else:
                st.error("Suba o template!")

    # ==========================================
    # LÓGICA: APRESENTAÇÃO COMERCIAL
    # ==========================================
    elif tipo_apresentacao == "Apresentação Comercial":
        with st.expander("💰 3. Detalhamento de Investimento e Parcelas", expanded=True):
            modo_logistica = st.radio("Como a Logística será tratada?", ["Estimada (Soma +30% automático)", "Cotada (Informar manualmente)"])
            
            st.write("**Fases de Investimento (Tabela):**")
            qtd_acoes = st.number_input("Quantas Macro Ações?", min_value=1, value=3)
            
            acoes_fin = []
            total_op1 = 0.0
            total_op2 = 0.0
            
            for i in range(qtd_acoes):
                cf1, cf2, cf3 = st.columns(3)
                n_acao = cf1.text_input(f"Ação {i+1}", key=f"ac_n_{i}")
                v1 = cf2.number_input(f"Valor Opção 1 (R$)", key=f"ac_v1_{i}", value=0.0)
                
                if modo_logistica == "Estimada (Soma +30% automático)":
                    v2 = v1 * 1.3
                    cf3.info(f"Opção 2: R$ {v2:,.2f}")
                else:
                    v2 = cf3.number_input(f"Valor Opção 2 (R$)", key=f"ac_v2_{i}", value=0.0)
                
                if n_acao:
                    acoes_fin.append({'nome': n_acao, 'v1': v1, 'v2': v2})
                    total_op1 += v1
                    total_op2 += v2
            
            st.markdown("---")
            st.write(f"**Total OP1:** R$ {total_op1:,.2f} | **Total OP2:** R$ {total_op2:,.2f}")
            
            st.write("**Condições de Pagamento:**")
            qtd_parcelas = st.number_input("Quantidade de Parcelas ({{QTD_PARCELAS}})", min_value=1, value=12)

        if st.button("🚀 GERAR PROPOSTA COMERCIAL"):
            if template_upload:
                mapa_comercial = {
                    "{{SERVICO}}": servico, "{{CLIENTE}}": cliente, "{{UNIDADE}}": unidade, 
                    "{{NUM_PROP}}": num_prop, "{{ESCOPO}}": escopo_tag,
                    "{{DATA}}": datetime.date.today().strftime("%d/%m/%Y"),
                    "{{JUSTIFICATIVA}}": justificativa, "{{OBJETIVO}}": objetivo,
                    "{{PRAZO}}": prazo, "{{FORMATO}}": formato, "{{IDIOMA}}": idioma, "{{IDAS}}": str(idas),
                    "{{VALOR_OP1}}": f"R$ {total_op1:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{VALOR_OP2}}": f"R$ {total_op2:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{VALOR_OP1_EXT}}": valor_por_extenso(total_op1),
                    "{{VALOR_OP2_EXT}}": valor_por_extenso(total_op2),
                    "{{QTD_PARCELAS}}": str(qtd_parcelas),
                    "{{VLR1_PARCELAS}}": f"R$ {(total_op1/qtd_parcelas):,.2f}".replace(",", "X").replace(".", ",").replace("X", "."),
                    "{{VLR2_PARCELAS}}": f"R$ {(total_op2/qtd_parcelas):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                }
                
                dist_parcelas = calcular_amortizacao(qtd_parcelas)
                dados_financeiros = {
                    'acoes': acoes_fin, 
                    'total_op1': total_op1, 
                    'total_op2': total_op2,
                    'parcelas': dist_parcelas,
                    'val_parcela': total_op2 / qtd_parcelas
                }
                
                pptx_io = processar_apresentacao(template_upload, mapa_comercial, atividades_lista, "Comercial", dados_financeiros)
                st.success("Comercial gerada com sucesso!")
                st.download_button("⬇️ Baixar", pptx_io, f"Comercial_{cliente}.pptx")
            else:
                st.error("Suba o template!")
