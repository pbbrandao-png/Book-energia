import streamlit as st
import pandas as pd
import unicodedata

# FUNÇÃO PARA PADRONIZAR COLUNAS
def limpar_coluna(texto):

    texto = str(texto).strip().upper()

    texto = unicodedata.normalize('NFKD', texto)\
        .encode('ASCII', 'ignore')\
        .decode('utf-8')

    return texto


# TÍTULO
st.title("Book de Energia")

# UPLOAD
arquivo = st.file_uploader(
    "Contratos aprovados",
    type=['xlsx', 'csv', 'xlsm']
)

# SE SUBIU ARQUIVO
if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")

    # LÊ A ABA
    df = pd.read_excel(
        arquivo,
        sheet_name='Contratos_Selecionados'
    )

    # PADRONIZA COLUNAS
    df.columns = [limpar_coluna(col) for col in df.columns]

    # MOSTRA SOMENTE AS COLUNAS
    st.dataframe(
        df[['CODIGO_WBC', 'OPERACAO', 'Tipo_Energia']],
        hide_index=True
    )
