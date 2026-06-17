"""
Geração do Recibo de Projeto (Handover) em PDF.

Recebe os MESMOS dicionários que a tela de Handover já monta
(dados_tabela e dados_logistica) e devolve os bytes de um PDF limpo,
paginado e sem nada cortado — pronto para um st.download_button.

Não captura a tela: monta o PDF a partir dos dados, então o resultado
não depende do navegador, do zoom nem do tamanho do monitor.
"""

import io
from datetime import datetime

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
)

# Paleta sóbria (ajuste à identidade visual da Comportamento se quiser)
COR_PRIMARIA = colors.HexColor("#1F3A5F")   # cabeçalho das tabelas / título
COR_LINHA = colors.HexColor("#D7DEE8")      # grade
COR_ZEBRA = colors.HexColor("#F2F5F9")      # linhas alternadas


def _estilos():
    base = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloRecibo", parent=base["Title"],
        fontName="Helvetica-Bold", fontSize=18, leading=22,
        textColor=COR_PRIMARIA, spaceAfter=2,
    )
    estilo_sub = ParagraphStyle(
        "SubRecibo", parent=base["Normal"],
        fontName="Helvetica", fontSize=10, leading=13,
        textColor=colors.HexColor("#555555"),
    )
    estilo_secao = ParagraphStyle(
        "SecaoRecibo", parent=base["Heading2"],
        fontName="Helvetica-Bold", fontSize=12, leading=15,
        textColor=COR_PRIMARIA, spaceBefore=14, spaceAfter=6,
    )
    estilo_campo = ParagraphStyle(
        "CelulaCampo", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=9, leading=12,
        textColor=colors.HexColor("#333333"),
    )
    estilo_valor = ParagraphStyle(
        "CelulaValor", parent=base["Normal"],
        fontName="Helvetica", fontSize=9, leading=12,
        textColor=colors.HexColor("#1A1A1A"),
    )
    estilo_header = ParagraphStyle(
        "CelulaHeader", parent=base["Normal"],
        fontName="Helvetica-Bold", fontSize=9, leading=12,
        textColor=colors.white,
    )
    return estilo_titulo, estilo_sub, estilo_secao, estilo_campo, estilo_valor, estilo_header


def _tabela_campo_valor(dados_tabela, larg_util, est_campo, est_valor, est_header):
    """Monta a tabela principal (Campo | Informação) com wrap de texto."""
    campos = dados_tabela.get("Campo", [])
    valores = dados_tabela.get("Informação", [])

    linhas = [[Paragraph("Campo", est_header), Paragraph("Informação", est_header)]]
    for campo, valor in zip(campos, valores):
        linhas.append([
            Paragraph(str(campo), est_campo),
            Paragraph(str(valor) if str(valor).strip() else "—", est_valor),
        ])

    larg_campo = 55 * mm
    tabela = Table(linhas, colWidths=[larg_campo, larg_util - larg_campo], repeatRows=1)
    estilo = [
        ("BACKGROUND", (0, 0), (-1, 0), COR_PRIMARIA),
        ("GRID", (0, 0), (-1, -1), 0.5, COR_LINHA),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for i in range(1, len(linhas)):
        if i % 2 == 0:
            estilo.append(("BACKGROUND", (0, i), (-1, i), COR_ZEBRA))
    tabela.setStyle(TableStyle(estilo))
    return tabela


def _tabela_logistica(dados_logistica, larg_util, est_campo, est_valor, est_header):
    """Monta a tabela de detalhamento logístico (cabeçalhos = chaves do dict)."""
    chaves = list(dados_logistica.keys())
    if not chaves:
        return None

    header = [Paragraph(str(k), est_header) for k in chaves]
    n_linhas = max(len(v) for v in dados_logistica.values())
    corpo = []
    for i in range(n_linhas):
        linha = []
        for k in chaves:
            col = dados_logistica[k]
            valor = col[i] if i < len(col) else ""
            linha.append(Paragraph(str(valor) if str(valor).strip() else "—", est_valor))
        corpo.append(linha)

    larg_col = larg_util / len(chaves)
    tabela = Table([header] + corpo, colWidths=[larg_col] * len(chaves), repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), COR_PRIMARIA),
        ("GRID", (0, 0), (-1, -1), 0.5, COR_LINHA),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return tabela


def gerar_pdf_handover(dados_tabela, dados_logistica, titulo="Recibo do Projeto",
                       subtitulo=""):
    """
    Devolve os bytes de um PDF (A4) com o recibo de handover.

    dados_tabela    -> {"Campo": [...], "Informação": [...]}
    dados_logistica -> {"<cabecalho 1>": [valor], "<cabecalho 2>": [valor], ...}
    titulo          -> título no topo do documento
    subtitulo       -> linha de contexto (ex.: cliente / nº da proposta)
    """
    buffer = io.BytesIO()
    margem = 18 * mm
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        leftMargin=margem, rightMargin=margem,
        topMargin=margem, bottomMargin=margem,
        title=titulo,
    )
    larg_util = A4[0] - 2 * margem

    est_titulo, est_sub, est_secao, est_campo, est_valor, est_header = _estilos()

    story = []
    story.append(Paragraph(titulo, est_titulo))
    if subtitulo:
        story.append(Paragraph(subtitulo, est_sub))
    story.append(Paragraph(
        f"Gerado em {datetime.now().strftime('%d/%m/%Y %H:%M')}", est_sub
    ))
    story.append(Spacer(1, 6))
    story.append(HRFlowable(width="100%", thickness=1, color=COR_PRIMARIA))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Informações Comerciais", est_secao))
    story.append(_tabela_campo_valor(dados_tabela, larg_util,
                                     est_campo, est_valor, est_header))

    tab_log = _tabela_logistica(dados_logistica, larg_util,
                                est_campo, est_valor, est_header)
    if tab_log is not None:
        story.append(Paragraph("Detalhamento Logístico", est_secao))
        story.append(tab_log)

    doc.build(story)
    return buffer.getvalue()
