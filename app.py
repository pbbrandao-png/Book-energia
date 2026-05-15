import streamlit as st

# Título do app
st.title("Book de Energia")

# Caixa de upload
arquivo = st.file_uploader(
    "Suba sua planilha",
    type=['xlsx', 'csv']
)

# Verifica se usuário subiu algo
if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")

    st.write("Nome do arquivo:")
    st.write(arquivo.name)
