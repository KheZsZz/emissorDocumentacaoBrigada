import streamlit as st
import pandas as pd
from io import BytesIO

def show():
    with st.sidebar:
        st.title("📁 Opções")
        st.subheader("⬇️ Baixar Arquivo Modelo")

        # Lista de colunas da sua planilha original
        colunas = [
            "nome_aluno",
            "cpf",
            "empresa",
            "cnpj",
            "endereco_empresa",
            "nivel_curso",
            "modalidade",
            "carga_horaria",
            "cidade_data",
            "instrutor",
            "documento_instrutor",
            "conclusao",
            "validade",
            "Foto"
        ]

        # Cria um DataFrame com valores exemplo
        exemplo = {col: [f"exemplo_{col}"] for col in colunas}

        df_modelo = pd.DataFrame(exemplo)

        # Exporta para Excel em memória
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_modelo.to_excel(writer, index=False, sheet_name="Modelo")
        buffer.seek(0)

        # Botão de download na barra lateral
        st.sidebar.download_button(
            label="📥 Baixar modelo Excel",
            data=buffer,
            file_name="modelo_planilha.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )