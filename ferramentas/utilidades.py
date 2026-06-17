import math
import streamlit as st
from ferramentas import config


def formatar_moeda(valor):
    """Valor literal com cifrão. Use em st.metric, st.table e nos PPTX."""
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def esc_md(texto):
    """
    Escapa o cifrão para uso em CONTEXTOS MARKDOWN (st.write/markdown/success/
    info/warning/caption). Sem isso, o Streamlit interpreta '$...$' como LaTeX
    e o cifrão some (e os valores entre eles saem em fonte de fórmula, menores).
    NÃO use em st.metric, st.table nem nos PPTX — lá o cifrão deve ser literal.
    """
    return texto.replace("$", r"\$")


# --- Conversão de valor por extenso ---
_UNIDADES = ["", "um", "dois", "três", "quatro", "cinco", "seis", "sete", "oito", "nove"]
_DEZ_A_DEZENOVE = ["dez", "onze", "doze", "treze", "quatorze", "quinze", "dezesseis",
                   "dezessete", "dezoito", "dezenove"]
_DEZENAS = ["", "", "vinte", "trinta", "quarenta", "cinquenta", "sessenta", "setenta",
            "oitenta", "noventa"]
_CENTENAS = ["", "cento", "duzentos", "trezentos", "quatrocentos", "quinhentos",
             "seiscentos", "setecentos", "oitocentos", "novecentos"]


def _converter_trio(n):
    """Converte um número de 0 a 999 para extenso."""
    if n == 100:
        return "cem"
    if n == 0:
        return ""
    c, resto = n // 100, n % 100
    d, u = resto // 10, resto % 10
    partes = []
    if c > 0:
        partes.append(_CENTENAS[c])
    if d == 1:
        partes.append(_DEZ_A_DEZENOVE[u])
    else:
        if d > 1:
            partes.append(_DEZENAS[d])
        if u > 0:
            partes.append(_UNIDADES[u])
    return " e ".join(partes)


def _converter_inteiro(inteiro):
    """Converte um inteiro (até bilhões) para extenso, com conector 'e' correto."""
    if inteiro == 0:
        return "zero"

    bilhoes = inteiro // 1_000_000_000
    milhoes = (inteiro % 1_000_000_000) // 1_000_000
    milhares = (inteiro % 1_000_000) // 1_000
    resto = inteiro % 1_000

    grupos = []
    if bilhoes > 0:
        grupos.append(_converter_trio(bilhoes) + (" bilhões" if bilhoes > 1 else " bilhão"))
    if milhoes > 0:
        grupos.append(_converter_trio(milhoes) + (" milhões" if milhoes > 1 else " milhão"))
    if milhares > 0:
        grupos.append("mil" if milhares == 1 else _converter_trio(milhares) + " mil")
    if resto > 0:
        grupos.append(_converter_trio(resto))

    if len(grupos) == 1:
        return grupos[0]
    if resto > 0 and (resto < 100 or resto % 100 == 0):
        return ", ".join(grupos[:-1]) + " e " + grupos[-1]
    return ", ".join(grupos)


def valor_por_extenso(valor):
    """Ex.: 47500.50 -> 'Quarenta e sete mil e quinhentos reais e cinquenta centavos'."""
    valor = round(float(valor), 2)
    inteiro = int(valor)
    centavos = int(round((valor - inteiro) * 100))

    if inteiro == 0 and centavos > 0:
        texto_cent = _converter_inteiro(centavos)
        unidade = "centavo" if centavos == 1 else "centavos"
        resultado = f"{texto_cent} {unidade}"
        return resultado[0].upper() + resultado[1:]

    texto_reais = _converter_inteiro(inteiro)
    moeda = "real" if inteiro == 1 else "reais"
    resultado = f"{texto_reais} {moeda}"

    if centavos > 0:
        texto_cent = _converter_inteiro(centavos)
        unidade = "centavo" if centavos == 1 else "centavos"
        resultado += f" e {texto_cent} {unidade}"

    return resultado[0].upper() + resultado[1:]


# --- Estatística e financeiro ---
def calcular_amostra(N, margem_erro, proporcao, z=config.Z_SCORE):
    if N <= 0:
        return 0
    numerador = N * (z ** 2) * proporcao * (1 - proporcao)
    denominador = (margem_erro ** 2) * (N - 1) + (z ** 2) * proporcao * (1 - proporcao)
    return math.ceil(numerador / denominador)


def calcular_amortizacao(qtd_parcelas):
    pesos_base = [20, 20, 10, 10] + [5] * 30
    pesos_ativos = pesos_base[:qtd_parcelas]
    soma = sum(pesos_ativos)
    percentuais = [round((p / soma) * 100) for p in pesos_ativos]
    diferenca = 100 - sum(percentuais)
    if diferenca != 0:
        percentuais[0] += diferenca
    return percentuais


# --- Controladores de estado (memória do app) ---
def ir_para(pagina):
    st.session_state.current_page = pagina
    st.session_state.tentou_gerar = False
    st.session_state.pptx_gerado = None


def adicionar_unidade():
    novo_id = len(st.session_state.unidades_dcs)
    st.session_state.unidades_dcs.append(
        {"id": novo_id, "nome": f"Nova Unidade {novo_id+1}", "pop_total": 0, "lideres": 0,
         "oac": 0, "visitas": 0, "aprofundamento": 0}
    )


def remover_unidade(id_remover):
    if len(st.session_state.unidades_dcs) > 1:
        st.session_state.unidades_dcs = [
            u for u in st.session_state.unidades_dcs if u['id'] != id_remover
        ]


def adicionar_fase():
    novo_id = max([f['id'] for f in st.session_state.fases_dcs], default=-1) + 1
    st.session_state.fases_dcs.append(
        {"id": novo_id, "nome": "Nova Etapa", "horas": 0, "presencial": False}
    )


def remover_fase(id_remover):
    st.session_state.fases_dcs = [
        f for f in st.session_state.fases_dcs if f['id'] != id_remover
    ]
