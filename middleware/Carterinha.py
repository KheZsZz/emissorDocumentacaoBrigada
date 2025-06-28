import streamlit as st
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase.pdfmetrics import stringWidth
import os
from io import BytesIO

def gerar_carteirinhas(df):
    largura_folha, altura_folha = A4
    largura_carteirinha = 9 * cm
    altura_carteirinha = 5.5 * cm

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=(largura_folha, altura_folha))

    x_inicio = 1 * cm
    y_inicio = altura_folha - 0.1 * cm

    lateral_superior = "public/lateral_superior.png"
    assinatura_path = "public/Instrutores.png"

    for index, row in df.iterrows():
        nome = row['nome_aluno']
        rg_cpf = row['cpf']
        curso = "Brigada de incêndio"
        nivel = row['nivel_curso']
        modalidade = row['modalidade']
        conclusão = row['conclusao']
        validade = row['validade']
        foto = row.get('Foto', None)

        # Frente (lado esquerdo)
        x_frente = x_inicio
        y_frente = y_inicio

        c.rect(x_frente, y_frente - altura_carteirinha, largura_carteirinha, altura_carteirinha)

        if os.path.exists(lateral_superior):
            largura_img = largura_carteirinha
            altura_img = 1 * cm
            x_img = x_frente
            y_img = y_frente - altura_img
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

        if pd.notna(foto) and os.path.exists(foto):
            c.drawImage(foto, x_frente + largura_carteirinha - 3*cm, y_frente - 4*cm, width=2.5*cm, height=3.5*cm, preserveAspectRatio=True)

        # Verso (lado direito)
        x_verso = x_inicio + largura_carteirinha + 0.5*cm
        y_verso = y_inicio

        c.rect(x_verso, y_verso - altura_carteirinha, largura_carteirinha, altura_carteirinha)

        if os.path.exists(lateral_superior):
            largura_img = largura_carteirinha
            altura_img = 1 * cm
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

        texto1 = "Certificamos que o portador indicado no verso concluiu com aproveitamento o curso:"
        fonte = "Helvetica-Bold"
        tamanho_fonte = 8
        largura_max = largura_carteirinha - 1*cm

        linhas = quebrar_texto(texto1, fonte, tamanho_fonte, largura_max)
        text_obj = c.beginText()
        text_obj.setFont(fonte, tamanho_fonte)
        text_obj.setTextOrigin(x_verso + 0.5*cm, y_verso - 1.5*cm)

        for linha in linhas:
            text_obj.textLine(linha)
        c.drawText(text_obj)

        texto2 = "Brigada de Incêndio - Prevenção e Combate a Incêndio e Primeiros Socorros."
        fonte = "Helvetica"
        tamanho_fonte = 8
        linhas = quebrar_texto(texto2, fonte, tamanho_fonte, largura_max)

        text_obj = c.beginText()
        text_obj.setFont(fonte, tamanho_fonte)
        text_obj.setTextOrigin(x_verso + 0.5*cm, y_verso - 2.2*cm)

        for linha in linhas:
            text_obj.textLine(linha)
        c.drawText(text_obj)

        c.drawString(x_verso + 0.5*cm, y_verso - 3*cm, "Conteúdo Programatico da tabela B.2 da IT-17.")
        c.drawString(x_verso + 0.5*cm, y_verso - 3.5*cm, "Nível: Intermediário - 08 Horas")

        if os.path.exists(assinatura_path):
            largura_assinatura = 4*cm
            altura_assinatura = 1.5*cm
            x_assinatura = x_verso + 0.5*cm
            y_assinatura = y_verso - 5*cm
            c.drawImage(assinatura_path, x_assinatura, y_assinatura, width=largura_assinatura, height=altura_assinatura, preserveAspectRatio=True)

        c.drawString(x_verso + 0.5*cm, y_verso - 4.4*cm, "________________________")
        c.drawString(x_verso + 0.5*cm, y_verso - 4.7*cm, "Instrutor:Fernando Henrique")
        c.drawString(x_verso + 0.5*cm, y_verso - 5*cm, "CPF: 026.476.979-14")

        c.drawString(x_verso + 5*cm, y_verso - 4.4*cm, "________________________")
        c.drawString(x_verso + 5*cm, y_verso - 4.7*cm, "Aluno")

        y_inicio -= altura_carteirinha + 0.5*cm
        if y_inicio < 3*cm:
            c.showPage()
            y_inicio = altura_folha - 3*cm

    c.save()
    buffer.seek(0)
    return buffer