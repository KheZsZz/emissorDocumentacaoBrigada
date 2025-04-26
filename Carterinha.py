import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import landscape, A4
from reportlab.lib.units import cm
import os

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
    rg_cpf = row['rg_cpf']
    curso = "Brigada de incêndio"
    nivel = row['nivel_curso']
    modalidade = row['modalidade']
    foto = row.get('Foto', None)

    # 🔵 DESENHAR A FRENTE (lado esquerdo)
    x_frente = x_inicio
    y_frente = y_inicio

    c.rect(x_frente, y_frente - altura_carteirinha, largura_carteirinha, altura_carteirinha)

    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_frente + 0.5*cm, y_frente - 1*cm, f"Nome: {nome}")

    c.setFont("Helvetica", 8)
    c.drawString(x_frente + 0.5*cm, y_frente - 2*cm, f"RG/CPF: {rg_cpf}")
    c.drawString(x_frente + 0.5*cm, y_frente - 3*cm, f"Curso: {curso}")
    c.drawString(x_frente + 0.5*cm, y_frente - 4*cm, f"Nível: {nivel}")
    c.drawString(x_frente + 0.5*cm, y_frente - 5*cm, f"Modalidade: {modalidade}")

    # Foto (opcional)
    if pd.notna(foto) and os.path.exists(foto):
        c.drawImage(foto, x_frente + largura_carteirinha - 3*cm, y_frente - 4*cm, width=2.5*cm, height=3.5*cm, preserveAspectRatio=True)

    # 🔵 DESENHAR O VERSO (lado direito)
    x_verso = x_inicio + largura_carteirinha + 0.5*cm  # espaço de 1 cm entre frente e verso
    y_verso = y_inicio

    c.rect(x_verso, y_verso - altura_carteirinha, largura_carteirinha, altura_carteirinha)

    c.setFont("Helvetica", 8)
    c.drawString(x_verso + 0.5*cm, y_verso - 1*cm, "Regulamento:")
    texto_verso = (
        "Esta carteirinha é pessoal e intransferível.\n"
        "Em caso de perda, comunicar a secretaria.\n"
        "Válido até: Dezembro/2025."
    )

    linhas = texto_verso.split("\n")
    y_texto = y_verso - 2*cm
    for linha in linhas:
        c.drawString(x_verso + 0.5*cm, y_texto, linha)
        y_texto -= 0.7*cm

    c.drawString(x_verso + 0.5*cm, y_verso - 5*cm, "Assinatura da Instituição:")

    # Atualizar posição para próxima carteirinha na folha
    y_inicio -= altura_carteirinha + 0.5*cm  # altura + espaçamento entre linhas

    # Se acabar o espaço na página, criar nova página
    if y_inicio < 3*cm:
        c.showPage()
        y_inicio = altura_folha - 3*cm

# 7. Salvar o PDF
c.save()

print("✅ Carteirinhas frente e verso lado a lado geradas!")
