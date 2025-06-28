import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import qrcode
from io import BytesIO
import os

# Função para gerar certificado
def gerar_certificado(dados, output_path):
    doc = SimpleDocTemplate(
        output_path, 
        pagesize=landscape(A4),
        rightMargin=2*cm, 
        leftMargin=4*cm,
        topMargin=2*cm, 
        bottomMargin=0*cm )
    
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle(
        name='TitleCenter',
        parent=styles['Title'],
        alignment=TA_CENTER,
        fontSize=48,
        spaceAfter=1    
    )
    
    style_normal = ParagraphStyle(
        name='NormalLeft',
        parent=styles['Normal'],
        alignment=TA_LEFT,
        fontSize=14,
        spaceAfter=30,
        leading=15
    )
    
    style_centered = ParagraphStyle(
        name='Centered',
        parent=styles['Normal'],
        alignment=TA_CENTER,
        fontSize=12,
        spaceAfter=12
    )

    elements = []
    
    # elements.append(Spacer(1, 50))
    elements.append(Paragraph("Certificado", style_title))
    elements.append(Spacer(1, 50))

    texto = f"""
    Nome do aluno: <b>{dados['nome_aluno']}</b><br/>
    RG/CPF: <b>{dados['rg_cpf']}</b><br/>
    Empresa: <b>{dados['empresa']}</b><br/> 
    Endereço: {dados['endereco_empresa']}<br/><br/>
    
    Certificamos que o aluno acima identificado realizou o treinamnto de "Brigada de Incêndio - 
    Prevenção e Combate a Incêndio e Primeiros Socorros", de acordo com as normas NBR14277 da 
    ABNT, e IT-17 do Corpo de Bombeiros, com o conteúdo programático da tabela B.2 da IT-17. <br/> <br/> 
    
    Nivel: <b>{dados['nivel_curso']}</b><br/> 
    Modalidade: <b>{dados['modalidade']}</b>.<br/>
    Carga horária: <b>{dados['carga_horaria']}</b> horas.<br/>
    """
    
    texto_cidade = f"""<b>{dados['cidade_data']}.</b><br/>""" 
    elements.append(Paragraph(texto, style_normal))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph(texto_cidade, style_centered))
    
    elements.append(Spacer(1, 5))

    # Dados para a assinatura
    assinaturas = [
        [
            f"""
            _____________________________
            {dados['instrutor']}
            Instrutor
            {dados['documento_instrutor']}""", 
            
            f"""
            _____________________________
            {dados['nome_aluno']}
            Aluno
            RG/CPF: {dados['rg_cpf']}
            """, 
            
            """
            _____________________________
            CRISTIANO REIS
            Responsavel técnico
            CPF: 214.135.358-01
            """
        ],
    ]

    # Criar tabela
    tabela_assinaturas = Table(assinaturas, colWidths=[7*cm, 7*cm, 7*cm])

    # Estilo da tabela
    tabela_assinaturas.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ]))

    # Adicionar no PDF
    elements.append(Spacer(1, 50))
    elements.append(tabela_assinaturas)

    # Função para desenhar a imagem diretamente no canvas
    def desenhar_imagem(canvas, doc):
        page_width, page_height = landscape(A4)
        imagem_path = "public/lateral_esquerda.png"
        x = -15  # posição x no papel (30 pontos da borda esquerda)
        y = 0  # posição y (ajustar conforme a necessidade)
        largura = 120  # largura da imagem em pontos
        altura = page_height   # altura da imagem em pontos
        canvas.drawImage(imagem_path, x, y, width=largura, height=altura, preserveAspectRatio=True)

    # Gerar QR Code
    # qr_texto = f"Certificado: {dados['nome_aluno']} - {dados['empresa']} - {dados['nivel_curso']}"
    # qr = qrcode.make(qr_texto)
    # buffer = BytesIO()
    # qr.save(buffer)
    # buffer.seek(0)

    # qr_img = Image(buffer, width=3*cm, height=3*cm)
    # qr_img.hAlign = 'CENTER'
    # elements.append(qr_img)
    
    #Geração do PDF
    doc.build(elements, onFirstPage=desenhar_imagem)

# -------------------------------
# Carregar planilha Excel
planilha = pd.read_excel('dados_alunos.xlsx')

# Pasta de saída
output_folder = "certificados"
os.makedirs(output_folder, exist_ok=True)

# Gerar certificados
for index, row in planilha.iterrows():
    dados = {
        "nome_aluno": row['nome_aluno'],
        "rg_cpf": row['rg_cpf'],
        "empresa": row['empresa'],
        "endereco_empresa": row['endereco_empresa'],
        "nivel_curso": row['nivel_curso'],
        "modalidade": row['modalidade'],
        "carga_horaria": str(row['carga_horaria']),
        "cidade_data": row['cidade_data'],
        "instrutor": row['instrutor'],
        "documento_instrutor": row['documento_instrutor'],
    }
    nome_arquivo = os.path.join(output_folder, f"certificado_{dados['nome_aluno'].replace(' ', '_')}.pdf")
    gerar_certificado(dados, nome_arquivo)
    print(f"Certificado gerado: {nome_arquivo}")

print("✅ Todos os certificados foram gerados a partir do Excel!")
