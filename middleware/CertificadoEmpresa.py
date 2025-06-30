import streamlit as st
import pandas as pd
from io import BytesIO
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.utils import ImageReader
import os

# Função para gerar certificado por empresa em um único arquivo PDF
def gerar_certificados_empresas(df):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4),
                            rightMargin=2*cm, leftMargin=4*cm,
                            topMargin=2*cm, bottomMargin=0*cm)

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle(name='TitleCenter', parent=styles['Title'], alignment=TA_CENTER, fontSize=48, spaceAfter=1)
    style_normal = ParagraphStyle(name='NormalLeft', parent=styles['Normal'], alignment=TA_LEFT, fontSize=14, spaceAfter=30, leading=15)
    style_centered = ParagraphStyle(name='Centered', parent=styles['Normal'], alignment=TA_CENTER, fontSize=12, spaceAfter=12)

    all_elements = []
    # empresas_unicas = df.drop_duplicates(subset=["empresa"])

    for _, dados in df.iterrows():
        elements = []
        elements.append(Paragraph("Certificado", style_title))
        elements.append(Spacer(1, 50))

        texto = f"""
        Empresa: <b>{dados['empresa']}</b><br/> 
        CNPJ: <b>{dados['cnpj']}</b><br/> 
        Endereço: <b>{dados['endereco_empresa']}</b><br/><br/>
        Certificamos que a empresa acima identificada promoveu o treinamento de “Brigada de Incêndio -
        Prevenção e Combate a Incêndio e Primeiros Socorros”, de acordo com as normas NBR 14277 da
        ABNT, e IT-17 do Corpo de Bombeiros. <br/> <br/> 
        Nível: <b>{dados['nivel_curso']}</b><br/> 
        Modalidade: <b>{dados['modalidade']}</b><br/>
        Carga horária: <b>{dados['carga_horaria']}</b><br/>
        """

        texto_cidade = """<b>São Paulo, 10 de Janeiro de 2025. <br/><br/>Treinamentos realizados de novembro/2024 a janeiro/2025.</b><br/>"""
        elements.append(Paragraph(texto, style_normal))
        elements.append(Spacer(1, 5))
        elements.append(Paragraph(texto_cidade, style_centered))
        elements.append(Spacer(1, 5))

        assinaturas = [[
            f"_____________________________\n{dados['instrutor']}\nInstrutor\nCPF:{dados['documento_instrutor']}",
            "_____________________________\nCRISTIANO REIS\nResponsável Técnico\nCPF: 214.135.358-01"
        ]]
        tabela_assinaturas = Table(assinaturas, colWidths=[10*cm, 10*cm])
        tabela_assinaturas.setStyle(TableStyle([
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ]))
        elements.append(Spacer(1, 50))
        elements.append(tabela_assinaturas)
        all_elements.extend(elements)
        all_elements.append(PageBreak())

    def desenhar_imagem(canvas, doc):
        lateral = "public/lateral_esquerda.png"
        instrutor_ass = "public/Instrutores.png"
        resp_tec_ass = "public/assinaturacristianoreis.png"
        canvas.drawImage(lateral, -15, 0, width=120, height=landscape(A4)[1])
        canvas.drawImage(instrutor_ass, 300, 130, width=120, height=80, preserveAspectRatio=True)
        canvas.drawImage(resp_tec_ass, 500, 130, width=120, height=80, preserveAspectRatio=True)

    doc.build(all_elements, onFirstPage=desenhar_imagem, onLaterPages=desenhar_imagem)
    buffer.seek(0)
    return buffer