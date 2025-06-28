import streamlit as st
import pandas as pd
from components import Menu
from middleware import CertificadoAluno, Carterinha

# Função com cache para ler o Excel
@st.cache_data
def carregar_planilha(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_")

    for col in df.columns:
        df[col] = df[col].astype(str)
    return df


def main():
    st.set_page_config(layout="wide")
    st.markdown("# 🏚️ Importação")
    # st.title ("Documentação")
    st.write("Criação de certificados, atestados e carterinhas.")

    arquivo = st.file_uploader(label="Importar arquivo modelo (excel)", type=["xlsx"])
    if arquivo is not None:
        try:
            df = carregar_planilha(arquivo)
            st.success("Planilha carregada com sucesso!")
            st.session_state["dataframe"] = df
            df.index = df.index + 1
            st.dataframe(df)

            left, rigth = st.columns(2)

            if "pdf_pronto" not in st.session_state:
                    st.session_state.carterinha = False
                    st.session_state.certificado = False

            with left:
                st.title("🏚️ Certificados")
                if "dataframe" in st.session_state:
                    df = st.session_state["dataframe"]
                    st.info(f"Total de registros: {len(df)}")
                else:
                    st.warning("Nenhuma planilha carregada ainda. Volte para a página de upload.")

                with st.spinner("Gerando certificados, por favor aguarde..."):
                    certificado_final = CertificadoAluno.gerar_certificados(df)
                    st.success("Certificados gerados!")
                    st.session_state.certificado = True

                    if st.session_state.certificado:
                        st.download_button("📥 Baixar PDF com todos os certificados", certificado_final, file_name="certificados_completos.pdf")


            with rigth:
                st.title("🏚️ Carterinhas")

                if "dataframe" in st.session_state:
                    df = st.session_state["dataframe"]
                    st.info(f"Total de registros: {len(df)}")
                else:
                    st.warning("Nenhuma planilha carregada ainda. Volte para a página de upload.")

                with st.spinner("Gerando carterinhas, por favor aguarde..."):
                    pdf_buffer = Carterinha.gerar_carteirinhas(df)
                    st.success("Carterinhas geradas!")
                    st.session_state.carterinha = True

                if st.session_state.carterinha:
                    st.download_button(label="📥 Baixar carterinhas",data=pdf_buffer,file_name="carteirinhas.pdf",mime="application/pdf")
                
                    

        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")


if __name__ == "__main__":
    main()