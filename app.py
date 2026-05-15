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


# FUNÇÃO PARA TRATAR TIPO ENERGIA
def tratar_tipo_energia(valor):

    mapa = {
        'Incentivada-50%': 'Incentivada-I5',
        'Incentivada-CQ50%': 'Incentivada-CQ5',
        'Incentivada-100%': 'Incentivada-I1',
        'Incentivada-0%': 'Incentivada-I0',
        
    }

    return mapa.get(valor, valor)


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

    # TRATA TIPO ENERGIA
    df['TIPO_ENERGIA'] = df['TIPO_ENERGIA'].apply(tratar_tipo_energia)

    # MOSTRA TABELA
    st.dataframe(
        df[['CODIGO_WBC', 'OPERACAO', 'TIPO_ENERGIA']],
        hide_index=True
    )
