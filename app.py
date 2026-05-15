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

    # LÊ A ABA CORRETA
    df = pd.read_excel(
        arquivo,
        sheet_name='Contratos_Selecionados'
    )

    # MOSTRA COLUNAS
    st.write("Colunas encontradas:")
    st.write(df.columns)

    # MOSTRA A COLUNA Codigo_WBC
    st.write("Codigo_WBC:")

    st.write(df['Codigo_WBC'])
