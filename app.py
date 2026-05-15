import streamlit as st
import pandas as pd


# Título do app
st.title("Book de Energia")

# Caixa de upload
arquivo = st.file_uploader(
    "Contratos aprovados",
    type=['xlsx', 'csv','xlsm']
)

# Verifica se usuário subiu algo
if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")
    
 # LER A PLANILHA
    df = pd.read_excel(arquivo)

    # MOSTRAR AS COLUNAS
    st.write(df.columns)

    # Mostra coluna Codigo_WBC
    st.write("Codigo_WBC:")

    st.write(df['Codigo_WBC'])
