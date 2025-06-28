import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import cm
import os
from reportlab.pdfbase.pdfmetrics import stringWidth

# 1. Carregar a planilha Excel
df = pd.read_excel("dados_alunos.xlsx")

# 2. Tamanho da folha A4 
largura_folha, altura_folha = A4

# 3. Tamanho da carteirinha
largura_carteirinha = 9 * cm
altura_carteirinha = 5.5 * cm

# 4. Criar a pasta de saída
output_folder = "carteirinhas"
os.makedirs(output_folder, exist_ok=True)

# 5. Criar o canvas para todas as carteirinhas juntas
pdf_path = os.path.join(output_folder, "carteirinhas_lado_a_lado.pdf")
c = canvas.Canvas(pdf_path, pagesize=(largura_folha, altura_folha))
# 6. Posições iniciais
x_inicio = 1 * cm  # margem esquerda
y_inicio = altura_folha - 0.1 * cm  # margem superior

for index, row in df.iterrows():
    nome = row['nome_aluno']
    rg_cpf = row['cpf']
    curso = "Brigada de incêndio"
    nivel = row['nivel_curso']
    modalidade = row['modalidade'],
    conclusão = row['conclusao']
    validade = row['validade']
    foto = row.get('Foto', None)

    # 🔵 DESENHAR A FRENTE (lado esquerdo)
    x_frente = x_inicio
    y_frente = y_inicio
    
    lateral_superior = "public/lateral_superior.png"  
      
    c.rect(x_frente, y_frente - altura_carteirinha, largura_carteirinha, altura_carteirinha)
    
     # Verifica se o arquivo existe e insere a imagem
    if os.path.exists(lateral_superior):
        largura_img = largura_carteirinha
        altura_img = 1 * cm
        x_img = x_frente
        y_img = y_frente - altura_img  # Começa no topo da carteirinha

        c.drawImage(lateral_superior, x_img, y_img, width=largura_img, height=altura_img, preserveAspectRatio=False)
    
    c.setFont("Helvetica", 8)
    c.drawString(x_frente + 0.5*cm, y_frente - 1.5*cm, f"Nome:")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_frente + 0.5*cm, y_frente - 2*cm, f"{nome}")

    c.setFont("Helvetica", 8)
    c.drawString(x_frente + 0.5*cm, y_frente - 2.5*cm, f"RG/CPF:")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_frente + 0.5*cm, y_frente - 3*cm, f"{rg_cpf}")
    
    c.setFont("Helvetica", 8)
    c.drawString(x_frente + 0.5*cm, y_frente - 3.5*cm, f"Conclusão:")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_frente + 0.5*cm, y_frente - 4*cm, f"{conclusão}")
    
    c.setFont("Helvetica", 8)
    c.drawString(x_frente + 0.5*cm, y_frente - 4.5*cm, f"Validade:")
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_frente + 0.5*cm, y_frente - 5*cm, f"{validade}")

    # Foto (opcional)
    if pd.notna(foto) and os.path.exists(foto):
        c.drawImage(foto, x_frente + largura_carteirinha - 3*cm, y_frente - 4*cm, width=2.5*cm, height=3.5*cm, preserveAspectRatio=True)
    
    # 🔵 DESENHAR O VERSO (lado direito)
    x_verso = x_inicio + largura_carteirinha + 0.5*cm  # espaço de 1 cm entre frente e verso
    y_verso = y_inicio

    c.rect(x_verso, y_verso - altura_carteirinha, largura_carteirinha, altura_carteirinha)

    x_verso = x_inicio + largura_carteirinha + 0.5*cm
    y_verso = y_inicio
    
    if os.path.exists(lateral_superior):
        largura_img = largura_carteirinha
        altura_img = 1 * cm  # altura da faixa superior
        x_img = x_verso
        y_img = y_verso - altura_img

        c.drawImage(lateral_superior, x_img, y_img, width=largura_img, height=altura_img, preserveAspectRatio=False)
       

    def quebrar_texto(texto, fonte, tamanho_fonte, largura_max):
        palavras = texto.split()
        linhas = []
        linha = ""
        
        for palavra in palavras:
            teste = f"{linha} {palavra}".strip()
            if stringWidth(teste, fonte, tamanho_fonte) <= largura_max:
                linha = teste
            else:
                linhas.append(linha)
                linha = palavra
        if linha:
            linhas.append(linha)
        return linhas

    # Texto a desenhar
    texto = "Certificamos que o portador indicado no verso concluiu com aproveitamento o curso:"
    fonte = "Helvetica-Bold"
    tamanho_fonte = 8
    largura_max = largura_carteirinha - 1*cm  # considerando 0.5cm de margem de cada lado

    # Quebra o texto em linhas
    linhas = quebrar_texto(texto, fonte, tamanho_fonte, largura_max)

    # Cria TextObject
    text_obj = c.beginText()
    text_obj.setFont(fonte, tamanho_fonte)
    text_obj.setTextOrigin(x_verso + 0.5*cm, y_verso - 1.5*cm)

    # Escreve linha por linha
    for linha in linhas:
        text_obj.textLine(linha)

    c.drawText(text_obj)

   
    c.setFont("Helvetica", 8)
    # Texto a desenhar
    texto = "Brigada de Incêndio - Prevenção e Combate a Incêndio e Primeiros Socorros."
    fonte = "Helvetica"
    tamanho_fonte = 8
    largura_max = largura_carteirinha - 1*cm  # considerando 0.5cm de margem de cada lado

    # Quebra o texto em linhas
    linhas = quebrar_texto(texto, fonte, tamanho_fonte, largura_max)

    # Cria TextObject
    text_obj = c.beginText()
    text_obj.setFont(fonte, tamanho_fonte)
    text_obj.setTextOrigin(x_verso + 0.5*cm, y_verso - 2.2*cm)

    # Escreve linha por linha
    for linha in linhas:
        text_obj.textLine(linha)

    c.drawText(text_obj)
   
    c.drawString(x_verso + 0.5*cm, y_verso - 3*cm, "Conteúdo Programatico da tabela B.2 da IT-17.")
    c.drawString(x_verso + 0.5*cm, y_verso - 3.5*cm, "Nível: Intermediário - 08 Horas")
    
    assinatura_path = "public/Instrutores.png"

    # Verifica se o arquivo existe e insere a imagem
    if os.path.exists(assinatura_path):
        largura_assinatura = 4*cm
        altura_assinatura = 1.5*cm
        x_assinatura = x_verso + 0.5*cm
        y_assinatura = y_verso - 5*cm  # Ajuste conforme o espaço disponível

    c.drawImage(assinatura_path, x_assinatura, y_assinatura, width=largura_assinatura, height=altura_assinatura, preserveAspectRatio=True) 


    c.drawString(x_verso + 0.5*cm, y_verso - 4.4*cm, "________________________")
    c.drawString(x_verso + 0.5*cm, y_verso - 4.7*cm, "Instrutor:Fernando Henrique")
    c.drawString(x_verso + 0.5*cm, y_verso - 5*cm, "CPF: 026.476.979-14")
    
    c.drawString(x_verso + 5*cm, y_verso - 4.4*cm, "________________________")
    c.drawString(x_verso + 5*cm, y_verso - 4.7*cm, "Aluno")
    
    # Atualizar posição para próxima carteirinha na folha
    y_inicio -= altura_carteirinha + 0.5*cm  # altura + espaçamento entre linhas

    # Se acabar o espaço na página, criar nova página
    if y_inicio < 3*cm:
        c.showPage()
        y_inicio = altura_folha - 3*cm

# 7. Salvar o PDF
c.save()

print("✅ Carteirinhas frente e verso lado a lado geradas!")
