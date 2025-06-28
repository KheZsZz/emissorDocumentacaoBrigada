import streamlit as st
from components import Menu

def main():
    st.markdown("# 🏚️ Importação")
    ct = Menu.show("Importação")

    st.title ("Documentação")
    st.write("Criação de certificados, atestados e carterinhas.")

    arquivo = st.file_uploader(label="Importar arquivo modelo (excel)")



if __name__ == "__main__":
    main()