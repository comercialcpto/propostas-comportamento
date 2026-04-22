import streamlit as st
import pandas as pd
import math
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import io
import datetime

# ==========================================
# 1. FUNÇÕES AUXILIARES E MATEMÁTICA
# ==========================================
VERDE_CPTO = RGBColor(0, 153, 116) 
CINZA_ESCURO = RGBColor(64, 64, 64)

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_celula_tabela(cell, texto):
    cell.text = str(texto)
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.name = "DIN Alternate"
            run.font.size = Pt(14)

def valor_por_extenso(valor):
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
    if milhoes > 0: resultado.append(converter_trio(milhoes) + (" milhões" if milhoes > 1 else " milhão"))
    if milhares > 0: resultado.append(converter_trio(milhares) + " mil")
    if resto > 0: resultado.append(converter_trio(resto))
        
    extenso_final = ", ".join(resultado).replace(", e", " e")
    return extenso_final.capitalize() + (" reais" if inteiro > 1 else " real")

def calcular_amostra(N, margem_erro, proporcao, z=1.96):
    if N <= 0: return 0
    numerador = N * (z**2) * proporcao * (1 - proporcao)
    denominador = (margem_erro**2) * (N - 1) + (z**2) * proporcao * (1 - proporcao)
    return math.ceil(numerador / denominador)

def calcular_amortizacao(qtd_parcelas):
    pesos_base = [20, 20, 10, 10] + [5] * 30
    pesos_ativos = pesos_base[:qtd_parcelas]
    soma = sum(pesos_ativos)
    percentuais = [round((p / soma) * 100) for p in pesos_ativos]
    diferenca = 100 - sum(percentuais)
    if diferenca != 0: percentuais[0] += diferenca
    return percentuais

# ==========================================
# 2. MOTOR DE PROCESSAMENTO PPTX
# ==========================================
def deletar_slide(prs, slide):
    id_dict = { s.id: [i, s.rId] for i, s in enumerate(prs.slides._sldIdLst) }
    prs.part.drop_rel(id_dict[slide.slide_id][1])
    del prs.slides._sldIdLst[id_dict[slide.slide_id][0]]

def remover_linha_tabela(table, row_idx):
    try:
        tr = table.rows[row_idx]._tr
        tr.getparent().remove(tr)
    except Exception: pass

def remover_coluna_tabela(table, col_idx):
    try:
        tbl = table._tbl
        grid = tbl.tblGrid
        col = grid.gridCol_lst[col_idx]
        grid.remove(col)
        for tr in tbl.tr_lst:
            tc = tr.tc_lst[col_idx]
            tr.remove(tc)
    except Exception: pass

def processar_apresentacao(template_file, mapa, atividades, tipo_doc, dados_fin=None, qtd_meses=12):
    prs = Presentation(template_file)
    slides_para_deletar = []

    for slide in prs.slides:
        deletar_este_slide = False
        
        for shape in slide.shapes:
            if hasattr(shape, "text") and "Para DCS" in shape.text:
                if mapa.get("{{SERVICO}}", "") != "Diagnóstico (DCS/Clima/DCMA)":
                    deletar_este_slide = True
                    break 

            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in mapa.items():
                            if key in run.text: run.text = run.text.replace(key, str(value))
            
            if shape.has_table:
                tbl = shape.table
                for row in tbl.rows:
                    for cell in row.cells:
                        if cell.text_frame:
                            for p in cell.text_frame.paragraphs:
                                for run in p.runs:
                                    for key, value in mapa.items():
                                        if key in run.text: run.text = run.text.replace(key, str(value))

                if len(tbl.columns) >= 12 and len(atividades) > 0:
                    colunas_para_deletar = list(range(qtd_meses + 1, len(tbl.columns)))
                    try:
                        largura_original = shape.width
                        largura_deletada = sum([tbl.columns[c].width for c in colunas_para_deletar])
                        nova_largura = largura_original - largura_deletada
                        shape.left = int((prs.slide_width - nova_largura) / 2)
                    except Exception: pass

                    for c_idx in reversed(colunas_para_deletar):
                        remover_coluna_tabela(tbl, c_idx)

                    linhas_deletar = list(range(len(atividades) + 1, len(tbl.rows)))
                    for r_idx in reversed(linhas_deletar):
                        remover_linha_tabela(tbl, r_idx)

                    for row_idx, atividade in enumerate(atividades):
                        target_row = row_idx + 1 
                        if target_row < len(tbl.rows):
                            row = tbl.rows[target_row]
                            cell = row.cells[0]
                            cell.text = atividade['nome']
                            
                            tamanho_str = len(atividade['nome'])
                            fonte_tamanho = 12
                            if tamanho_str > 60: fonte_tamanho = 8
                            elif tamanho_str > 40: fonte_tamanho = 9
                            elif tamanho_str > 20: fonte_tamanho = 10

                            if cell.text_frame.paragraphs:
                                p = cell.text_frame.paragraphs[0]
                                if p.runs:
                                    run = p.runs[0]
                                    run.font.name = "Calibri"
                                    run.font.size = Pt(fonte_tamanho)
                                    run.font.color.rgb = CINZA_ESCURO

                            for m_idx in range(1, len(tbl.columns)):
                                if m_idx in atividade['meses']:
                                    cell_mes = row.cells[m_idx]
                                    cell_mes.fill.solid()
                                    cell_mes.fill.fore_color.rgb = VERDE_CPTO

                if tipo_doc == "Comercial" and dados_fin:
                    try:
                        cabecalho = tbl.rows[0].cells[0].text.strip().lower()
                        
                        if "macro" in cabecalho:
                            acoes = dados_fin['acoes']
                            linhas_para_deletar = []
                            for idx in range(1, len(tbl.rows)):
                                cell_text = tbl.rows[idx].cells[0].text.strip().lower()
                                if "investimento total" in cell_text:
                                    formatar_celula_tabela(tbl.rows[idx].cells[1], formatar_moeda(dados_fin['total_op1']))
                                    formatar_celula_tabela(tbl.rows[idx].cells[2], formatar_moeda(dados_fin['total_op2']))
                                elif idx <= len(acoes):
                                    formatar_celula_tabela(tbl.rows[idx].cells[0], acoes[idx-1]['nome'])
                                    formatar_celula_tabela(tbl.rows[idx].cells[1], formatar_moeda(acoes[idx-1]['v1']))
                                    formatar_celula_tabela(tbl.rows[idx].cells[2], formatar_moeda(acoes[idx-1]['v2']))
                                else:
                                    linhas_para_deletar.append(idx)
                            for idx in reversed(linhas_para_deletar):
                                remover_linha_tabela(tbl, idx)

                        elif "meses" in cabecalho:
                            parcelas = dados_fin['parcelas']
                            linhas_para_deletar = []
                            for idx in range(1, len(tbl.rows)):
                                cell_text = tbl.rows[idx].cells[0].text.strip().lower()
                                if "total" in cell_text and "investimento" not in cell_text:
                                    formatar_celula_tabela(tbl.rows[idx].cells[1], "100%")
                                    formatar_celula_tabela(tbl.rows[idx].cells[2], formatar_moeda(dados_fin['total_op2']))
                                elif idx <= len(parcelas):
                                    formatar_celula_tabela(tbl.rows[idx].cells[0], f"M{idx}")
                                    formatar_celula_tabela(tbl.rows[idx].cells[1], f"{parcelas[idx-1]}%")
                                    val_calc = dados_fin['total_op2'] * (parcelas[idx-1] / 100)
                                    formatar_celula_tabela(tbl.rows[idx].cells[2], formatar_moeda(val_calc))
                                else:
                                    linhas_para_deletar.append(idx)
                            for idx in reversed(linhas_para_deletar):
                                remover_linha_tabela(tbl, idx)
                    except Exception: pass 

        if deletar_este_slide: slides_para_deletar.append(slide)

    for slide in slides_para_deletar: deletar_slide(prs, slide)

    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# ==========================================
# 3. INTERFACE DE USUÁRIO E ESTADOS (WIZARD)
# ==========================================
st.set_page_config(page_title="Sistema Comportamento", layout="wide")

if "pptx_gerado" not in st.session_state: st.session_state.pptx_gerado = None
if "nome_arquivo" not in st.session_state: st.session_state.nome_arquivo = ""
if 'tentou_gerar' not in st.session_state: st.session_state.tentou_gerar = False

if 'current_page' not in st.session_state: st.session_state.current_page = "🏠 Início"
def ir_para(pagina):
    st.session_state.current_page = pagina
    st.session_state.tentou_gerar = False
    st.session_state.pptx_gerado = None

# Variáveis globais de inteligência
if 'unidades_dcs' not in st.session_state: st.session_state.unidades_dcs = [{"id": 0, "nome": "Unidade Matriz", "pop_total": 0, "lideres": 0}]
if 'fases_dcs' not in st.session_state:
    st.session_state.fases_dcs = [
        {"id": 0, "nome": "Planejamento / Reunião de Abertura", "horas": 4},
        {"id": 1, "nome": "Workshop - Gestão de Cultura de Segurança", "horas": 2},
        {"id": 2, "nome": "Elaboração do Relatório", "horas": 56},
        {"id": 3, "nome": "Apresentação dos Resultados", "horas": 4},
        {"id": 4, "nome": "Plano de Transformação Cultural (PTC)", "horas": 12},
        {"id": 5, "nome": "Suporte e Acompanhamento", "horas": 64}
    ]
if 'memoria_geral' not in st.session_state:
    st.session_state.memoria_geral = {
        "cliente": "", "unidade": "", "num_prop": "", "escopo": "", "prazo": "", 
        "formato": "Híbrido", "justificativa": "", "objetivo": "", "idas": 0
    }
if 'valores_finais' not in st.session_state:
    st.session_state.valores_finais = {"op1": 0.0, "op2": 0.0}

def adicionar_unidade():
    novo_id = len(st.session_state.unidades_dcs)
    st.session_state.unidades_dcs.append({"id": novo_id, "nome": f"Nova Unidade {novo_id+1}", "pop_total": 0, "lideres": 0})

def remover_unidade(id_remover):
    if len(st.session_state.unidades_dcs) > 1:
        st.session_state.unidades_dcs = [u for u in st.session_state.unidades_dcs if u['id'] != id_remover]

def adicionar_fase():
    novo_id = max([f['id'] for f in st.session_state.fases_dcs], default=-1) + 1
    st.session_state.fases_dcs.append({"id": novo_id, "nome": "Nova Etapa", "horas": 0})

def remover_fase(id_remover):
    st.session_state.fases_dcs = [f for f in st.session_state.fases_dcs if f['id'] != id_remover]


st.sidebar.title("🧭 Navegação Integrada")
menu_options = ["🏠 Início", "💰 1. Precificação", "📝 2. Proposta Técnica", "📈 3. Proposta Comercial"]
selecao_menu = st.sidebar.radio("Etapa do Projeto:", menu_options, index=menu_options.index(st.session_state.current_page))
st.session_state.current_page = selecao_menu

# ==========================================
# MÓDULO 1: INÍCIO
# ==========================================
if st.session_state.current_page == "🏠 Início":
    st.title("Bem-vindo ao Sistema Integrado - Comportamento")
    st.write("Siga o fluxo integrado: calcule os custos, gere a proposta técnica e, em seguida, a comercial. Os dados acompanham você!")
    st.button("🚀 Iniciar Novo Projeto (Ir para Precificação)", on_click=ir_para, args=("💰 1. Precificação",))

# ==========================================
# MÓDULO 2: PRECIFICAÇÃO
# ==========================================
elif st.session_state.current_page == "💰 1. Precificação":
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
        
        for und in st.session_state.unidades_dcs:
            if und['pop_total'] > 0:
                amostra_total = calcular_amostra(und['pop_total'], margem_erro=0.08, proporcao=0.5)
                amostra_lideres = calcular_amostra(und['lideres'], margem_erro=0.05, proporcao=0.8)
                
                turmas_hm = math.ceil(amostra_lideres / 12)
                turmas_foco = math.ceil((amostra_total - amostra_lideres) / 12) if amostra_total > amostra_lideres else 0
                horas_hm = turmas_hm * 2
                horas_foco = turmas_foco * 2
                entrevistas = max(0, amostra_lideres - (turmas_hm * 12)) if amostra_lideres > 0 else 8
                horas_entrevistas = entrevistas * 1.5
                
                total_horas_unidade = horas_hm + horas_foco + horas_entrevistas
                total_horas_campo += total_horas_unidade
                total_amostra += amostra_total

        st.info(f"O motor estatístico gerou um total de **{total_horas_campo} horas de campo** com base nas unidades informadas.")

        st.markdown("---")
        st.markdown("### 📋 3. Plano Detalhado (Etapas Adicionais e Backoffice)")
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

        st.markdown("---")
        st.markdown("### 💰 4. Precificação Final e Logística")
        
        horas_totais = total_horas_campo + total_horas_fases
        valor_op1 = horas_totais * taxa_hora
        
        tipo_logistica = st.selectbox("Formato de cálculo de deslocamento:", [
            "1. Sem Logística (100% pelo Cliente)", 
            "2. Logística Estimada (Percentual %)"
        ])
        
        logistica_total = 0.0
        if tipo_logistica == "2. Logística Estimada (Percentual %)":
            perc_logistica = st.number_input("Margem Estimada de Logística (%)", min_value=0, max_value=100, value=30)
            logistica_total = valor_op1 * (perc_logistica / 100)
        
        valor_op2 = valor_op1 + logistica_total
        
        st.session_state.valores_finais["op1"] = valor_op1
        st.session_state.valores_finais["op2"] = valor_op2
        
        st.success(f"**Carga Horária Total:** {horas_totais} horas")
        c_tot1, c_tot2 = st.columns(2)
        c_tot1.metric("Total OP1 (Serviço Técnico)", formatar_moeda(valor_op1))
        c_tot2.metric("Total OP2 (Com Logística)", formatar_moeda(valor_op2))

        st.write("")
        st.button("Salvar e Avançar para Proposta Técnica ➡️", on_click=ir_para, args=("📝 2. Proposta Técnica",), type="primary")

# ==========================================
# MÓDULO 3: PROPOSTA TÉCNICA
# ==========================================
elif st.session_state.current_page == "📝 2. Proposta Técnica":
    st.title("📝 2. Gerador de Proposta Técnica")
    st.caption("As informações inseridas aqui alimentarão automaticamente a Proposta Comercial.")
    
    with st.sidebar:
        st.markdown("---")
        template_upload = st.file_uploader(f"Suba o template (Técnica)", type="pptx")
    
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
        # Importa do session_state da precificação + a coleta que é implícita
        fases_importadas = [{"nome": "Coleta de Dados"}] + st.session_state.fases_dcs
        
        for i, fase in enumerate(fases_importadas):
            ca, cm = st.columns([0.4, 0.6])
            nome_at = ca.text_input(f"Nome da Fase {i+1}", value=fase['nome'], key=f"tg_{i}")
            meses_at = cm.multiselect("Selecione os meses", list(range(1, int(qtd_meses_projeto) + 1)), key=f"tm_{i}")
            if meses_at:
                atividades_lista.append({"nome": nome_at, "meses": meses_at})

    with st.expander("👥 3. Detalhamento Simplificado do Público", expanded=True):
        st.info("Os totais foram carregados da Precificação. Ajuste conforme necessário.")
        
        lideres_calc = sum(u['lideres'] for u in st.session_state.unidades_dcs)
        pop_calc = sum(u['pop_total'] for u in st.session_state.unidades_dcs)
        
        cp1, cp2, cp3 = st.columns(3)
        n_lid_total = cp1.number_input("Total de Líderes", min_value=0, value=int(lideres_calc))
        n_oper = cp2.number_input("Total de Operacionais / Base", min_value=0, value=int(max(0, pop_calc - lideres_calc)))
        n_terc = cp3.number_input("Terceiros (Adicionais)", min_value=0, value=0)
        
        n_p_terc = n_lid_total + n_oper + n_terc
        st.metric("Total do Público Alvo", n_p_terc)

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
                # As macros de detalhamento foram resumidas
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

# ==========================================
# MÓDULO 4: PROPOSTA COMERCIAL
# ==========================================
elif st.session_state.current_page == "📈 3. Proposta Comercial":
    st.title("📈 3. Gerador de Proposta Comercial")
    st.caption("A mágica acontece aqui: não é preciso digitar mais nada além das parcelas.")
    
    with st.sidebar:
        st.markdown("---")
        template_upload = st.file_uploader(f"Suba o template (Comercial)", type="pptx")

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
