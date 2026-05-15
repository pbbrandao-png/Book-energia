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
    
# FUNÇÃO PARA TRATAR SUBMERCADO
def tratar_submercado(valor):

    mapa = {
        'N': 'NORTE',
        'Sul': 'SUL',
        'NE': 'NORDESTE',
        'SE/CO': 'SUDESTE',
    }

    return mapa.get(valor, valor)

# FORMATA MWH
if 'MONTANTE MWh' in df.columns:

    df['MONTANTE MWh'] = df['MONTANTE MWh'].apply(
        lambda x: f"{x:,.3f}"
        .replace(',', 'X')
        .replace('.', ',')
        .replace('X', '.')
    )

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
            'CODIGO_WBC': 'BOLETA',
            'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
            'QUANTATUALIZADA': 'MONTANTE MWh'
        }
    )

    # TRATA FONTE
    if 'FONTE' in df.columns:

        df['FONTE'] = df['FONTE'].apply(tratar_fonte)

    # TRATA SUBMERCADO
    if 'SUBMERCADO' in df.columns:

        df['SUBMERCADO'] = df['SUBMERCADO'].apply(tratar_submercado)
    
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

    # COLUNAS
    colunas_desejadas = [
        'BOLETA',
        'OPERACAO',
        'FONTE',
        'PARTE',
        'CONTRAPARTE',
        'CP/LP',
        'SUBMERCADO',
        'MONTANTE MWh'
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
