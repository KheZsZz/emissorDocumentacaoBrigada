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
    Empresa: <b>{dados['empresa']}</b><br/> 
    CNPJ: <b>{dados['cnpj']}</b><br/> 
    Endereço: <b>{dados['endereco_empresa']}</b><br/><br/>
    
    Certificamos que a empresa acima identificada promoveu o treinamento de “Brigada de Incêndio -
    Prevenção e Combate a Incêndio e Primeiros Socorros”, de acordo com as normas NBR 14277 da
    ABNT, e IT-17 do Corpo de Bombeiros. <br/> <br/> 
    
    Nivel: <b>{dados['nivel_curso']}</b><br/> 
    Modalidade: <b>{dados['modalidade']}</b><br/>
    Carga horária: <b>{dados['carga_horaria']}</b><br/>
    """
    
    texto_cidade = f"""<b>São Paulo, 10 de Janeiro de 2025. <br/><br/>Treinamentos realizados de novembro/2024 a janeiro/2025.</b><br/>""" 
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
            CPF:{dados['documento_instrutor']}""", 
            
            f"""
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
        
        #assinaturas
        assinatura_Resp_tec = "public/assinaturacristianoreis.png"
        x = 500
        y = 130
        altura = 80
        canvas.drawImage(assinatura_Resp_tec, x, y, width=largura, height=altura, preserveAspectRatio=True)
        
        assinaura_instrutor = "public/Instrutores.png"
        x = 300  
        y = 130
        altura = 80
        canvas.drawImage(assinaura_instrutor, x, y, width=largura, height=altura, preserveAspectRatio=True)
        
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
output_folder = "certificadosEmpresa"
os.makedirs(output_folder, exist_ok=True)

# Gerar certificados
for index, row in planilha.iterrows():
    dados = {
        "empresa": row['empresa'],
        "cnpj":row['cnpj'],
        "endereco_empresa": row['endereco_empresa'],
        "nivel_curso": row['nivel_curso'],
        "modalidade": row['modalidade'],
        "carga_horaria": str(row['carga_horaria']),
        "cidade_data": row['cidade_data'],
        "instrutor": row['instrutor'],
        "documento_instrutor": row['documento_instrutor'],
    }
    
    nome_arquivo = os.path.join(
        output_folder,
        f"certificado_{dados['cnpj'].replace('/', '-').replace('.', '').replace(' ', '')}.pdf"
    )
    
    gerar_certificado(dados, nome_arquivo)
    
    print(f"Certificado gerado: {nome_arquivo}")

print("✅ Todos os certificados foram gerados a partir do Excel!")
