import streamlit as st
import pandas as pd

# Título
st.title("Book de Energia")

# Upload
arquivo = st.file_uploader(
    "Contratos aprovados",
    type=['xlsx', 'csv', 'xlsm']
)

# Se subiu arquivo
if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")

    # Lê a aba correta
    df = pd.read_excel(
        arquivo,
        sheet_name='Contratos_Selecionados'
    )

    # Mostra apenas a coluna Codigo_WBC
    st.dataframe(
        df[['Codigo_WBC']],
        hide_index=True
    )
