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


# FUNÇÃO PARA TRATAR FONTE
def tratar_fonte(valor):

    mapa = {
        'Incentivada 50%': 'Incentivada-I5',
        'Cogeração Qualificada 50%': 'Incentivada-CQ5',
        'Incentivada 100%': 'Incentivada-I1',
        'Incentivada 0%': 'Incentivada-I0',
    }

    return mapa.get(valor, valor)


# TÍTULO
st.title("Livro de Energia - Abril/2026")

# UPLOAD
arquivo = st.file_uploader(
    "Contratos aprovados",
    type=['xlsx', 'csv', 'xlsm']
)

# SE SUBIU ARQUIVO
if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")

    # LÊ O EXCEL
    df = pd.read_excel(
        arquivo,
        skiprows=8
    )

    # PADRONIZA COLUNAS
    df.columns = [limpar_coluna(col) for col in df.columns]

    # MOSTRA COLUNAS EXISTENTES
    st.write("Colunas encontradas no arquivo:")
    st.write(df.columns.tolist())

    # RENOMEIA COLUNAS
    df = df.rename(
        columns={
            'PARTE_NOME_FANTASIA': 'PARTE',
            'MOVIMENTACAO': 'OPERACAO',
            'FONTE_CONTRATO': 'FONTE',
            'CODIGO_WBC': 'BOLETA'
        }
    )

    # TRATA FONTE
    if 'FONTE' in df.columns:

        df['FONTE'] = df['FONTE'].apply(tratar_fonte)

    # CONVERTE DATAS
    df['SUPRIMENTO_INICIO'] = pd.to_datetime(
        df['SUPRIMENTO_INICIO'],
        errors='coerce'
    )

    df['SUPRIMENTO_TERMINO'] = pd.to_datetime(
        df['SUPRIMENTO_TERMINO'],
        errors='coerce'
    )

    # CALCULA DIFERENÇA EM DIAS
    df['DIAS'] = (
        df['SUPRIMENTO_TERMINO']
        - df['SUPRIMENTO_INICIO']
    ).dt.days

    # CRIA COLUNA CP/LP
    df['CP/LP'] = df['DIAS'].apply(
        lambda x: 'LP' if x > 31 else 'CP'
    )

    # COLUNAS QUE VOCÊ QUER MOSTRAR
    colunas_desejadas = [
        'BOLETA',
        'OPERACAO',
        'FONTE',
        'PARTE',
        'CP/LP'
    ]

    # VERIFICA QUAIS EXISTEM
    colunas_existentes = [
        col for col in colunas_desejadas
        if col in df.columns
    ]

    # MOSTRA TABELA
    st.dataframe(
        df[colunas_existentes],
        hide_index=True
    )
