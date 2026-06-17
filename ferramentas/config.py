"""
Parâmetros centrais do Sistema Comportamento.
Todos os "números mágicos" de negócio ficam aqui. Para ajustar uma regra
(ex.: preço do combustível, margem de erro estatística, taxa de logística),
altere SÓ neste arquivo e o sistema inteiro acompanha.
"""
# --- Estatística (Motor DCS) ---
Z_SCORE = 1.96               # nível de confiança ~95%
MARGEM_ERRO_GERAL = 0.08     # 8% para a população geral
PROPORCAO_GERAL = 0.50       # heterogeneidade máxima (pior caso)
MARGEM_ERRO_LIDERES = 0.05   # 5% para a liderança (mais rigoroso)
PROPORCAO_LIDERES = 0.80     # homogeneidade esperada na liderança
TAMANHO_TURMA = 12           # pessoas por turma / grupo focal
HORAS_POR_TURMA = 2          # h por turma de Hearts & Minds e por grupo focal
HORAS_POR_ENTREVISTA = 1.5   # h por entrevista individual

# --- Atividades manuais de campo (DCS) — somam às horas presenciais ---
HORAS_POR_OAC = 2            # h por OAC
HORAS_POR_VISITA = 1         # h por visita técnica na área
HORAS_POR_APROFUNDAMENTO = 1 # h por atividade de aprofundamento

# --- Precificação ---
TAXA_HORA_PADRAO = 480.0     # valor sugerido da hora técnica
# --- Logística ---
TAXI_BASE_IDA = 150.0            # valor de um trajeto de táxi (x2 = ida e volta)
CUSTO_ALIMENTACAO_DIA = 120.0    # alimentação por dia/ida
COMBUSTIVEL_KM_POR_LITRO = 9.0   # autonomia média do carro
COMBUSTIVEL_PRECO_LITRO = 6.0    # preço do litro
AEREO_FATOR = 3.2                # heurística da média ponderada do aéreo
AEREO_DIVISOR = 6.0
TAXA_LOGISTICA = 1.10            # +10% de tributação sobre aéreo e carro
# Defaults por módulo (a logística é compartilhada, mas os padrões mudam)
PERCENTUAL_LOG_DCS = 30          # default da logística estimada (Diagnóstico)
PERCENTUAL_LOG_PONTUAL = 15      # default da logística estimada (Pontual)
DIAS_IDA_DCS = 5                 # semana cheia
DIAS_IDA_PONTUAL = 1             # palestra normalmente 1 dia
