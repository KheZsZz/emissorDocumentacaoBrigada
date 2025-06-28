import streamlit as st

def show(page_name):
    with st.sidebar:
        st.title("Menu")
        
        if page_name == "Importação":
            options = ["Selecione um centro...", "Rochacara", "Treinnar"]
            ct = st.selectbox("Centros de treinamentos", options, index=0)    

    if ct == "Selecione um centro...":
        st.warning("Por favor, selecione um centro de treinamento.")
        return None

    return ct