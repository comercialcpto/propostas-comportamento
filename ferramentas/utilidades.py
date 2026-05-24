import math
import streamlit as st

def formatar_moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

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

# --- CONTROLADORES DE ESTADO (MEMÓRIA DO APP) ---
def ir_para(pagina):
    st.session_state.current_page = pagina
    st.session_state.tentou_gerar = False
    st.session_state.pptx_gerado = None

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
