# =============================================================================
#
# ⚡ BOOK DE ENERGIA
#
# =============================================================================

import streamlit as st
import pandas as pd
import unicodedata
import calendar


# =============================================================================
# CONFIGURAÇÕES
# =============================================================================

TIPOS_ARQUIVO = ['xlsx', 'xlsm', 'csv']

COLUNAS_EXIBICAO = [
    'BOLETA',
    'OPERACAO',
    'FONTE',
    'PARTE',
    'CONTRAPARTE',
    'CP/LP',
    'CONTRAPARTE_CNPJ',
    'SUBMERCADO',
    'MONTANTE MWh',
    'MONTANTE MWm',
    'CLIQ PARADIGMA',
    'Cliq Mês Anterior',
    'MODULACAO WBC',
    'MOD MIN',
    'MOD MAX',
]


# =============================================================================
# FUNÇÕES DE TEXTO
# =============================================================================

def limpar_coluna(texto):
    """Remove espaços, acentos e deixa em maiúsculo."""

    texto = str(texto).strip().upper()

    texto = (
        unicodedata
        .normalize('NFKD', texto)
        .encode('ASCII', 'ignore')
        .decode('utf-8')
    )

    return texto


def formatar_numero_br(valor, casas=2):
    """Formata número no padrão brasileiro (ex: 1.234,56)."""

    try:
        return (
            f"{valor:,.{casas}f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )
    except Exception:
        return '-'
        
def formatar_cnpj(valor):
    """Formata número no padrão de CNPJ (ex: 12.345.678/0001-99)."""

    try:
        digits = ''.join(filter(str.isdigit, str(valor)))
        if len(digits) != 14:
            return str(valor)
        return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:]}"
    except Exception:
        return str(valor)


# =============================================================================
# FUNÇÕES DE NEGÓCIO
# =============================================================================

def classificar_cp_lp(dias):
    """Retorna 'CP' ou 'LP' com base na quantidade de dias."""

    if pd.isna(dias):
        return '-'

    return 'LP' if dias > 31 else 'CP'


def total_horas_mes(mes, ano):
    """Retorna o total de horas de um determinado mês/ano."""

    try:
        dias_mes = calendar.monthrange(int(ano), int(mes))[1]
        return dias_mes * 24
    except Exception:
        return None

def tratar_cnpj(df):
    if 'CONTRAPARTE_CNPJ' not in df.columns:
        return df

    df['CONTRAPARTE_CNPJ'] = (
        df['CONTRAPARTE_CNPJ']
        .astype(str)
        .str.strip()
        .apply(formatar_cnpj)
    )

    return df

df = tratar_cnpj(df)


# =============================================================================
# UPLOADS
# =============================================================================

st.set_page_config(page_title='Book de Energia', layout='wide')
st.title('⚡ Book de Energia')

col1, col2 = st.columns(2)

with col1:
    arquivo_aprovados = st.file_uploader(
        label='Contratos aprovados',
        type=TIPOS_ARQUIVO,
        key='aprovados'
    )

with col2:
    arquivo_mes_anterior = st.file_uploader(
        label='Contratos mês anterior',
        type=TIPOS_ARQUIVO,
        key='mes_anterior'
    )

if arquivo_aprovados is None:
    st.info('📂 Faça upload do arquivo principal.')
    st.stop()


# =============================================================================
# LEITURA DO ARQUIVO PRINCIPAL
# =============================================================================

try:
    df = pd.read_excel(arquivo_aprovados, skiprows=8)
    st.success('✅ Arquivo principal carregado!')
except Exception as erro:
    st.error(f'❌ Erro ao ler arquivo principal: {erro}')
    st.stop()


# =============================================================================
# LIMPEZA DOS NOMES DAS COLUNAS
# =============================================================================

def limpar_nomes_colunas(df):
    df.columns = [limpar_coluna(col) for col in df.columns]
    return df

df = limpar_nomes_colunas(df)


# =============================================================================
# RENOMEAÇÃO DAS COLUNAS
# =============================================================================

def renomear_colunas(df):
    df = df.rename(columns={
        'PARTE_NOME_FANTASIA':       'PARTE',
        'MOVIMENTACAO':              'OPERACAO',
        'FONTE_CONTRATO':            'FONTE',
        'CODIGO_WBC':                'BOLETA',
        'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
        'QUANTATUALIZADA':           'MONTANTE_MWH',
        'CODIGO_CCEE':               'CLIQ PARADIGMA',
        'TIPO_DE_MODULACAO':         'MODULACAO WBC',
        'FLEXLIMITE_MODULACAOMAX':   'MOD MAX',
        'FLEXLIMITE_MODULACAOMIN':   'MOD MIN',
    })
    return df

df = renomear_colunas(df)


# =============================================================================
# DATAS DE SUPRIMENTO
# =============================================================================

def converter_datas(df):
    if 'SUPRIMENTO_INICIO' not in df.columns or 'SUPRIMENTO_TERMINO' not in df.columns:
        return df

    df['SUPRIMENTO_INICIO']  = pd.to_datetime(df['SUPRIMENTO_INICIO'],  errors='coerce')
    df['SUPRIMENTO_TERMINO'] = pd.to_datetime(df['SUPRIMENTO_TERMINO'], errors='coerce')

    return df

df = converter_datas(df)


# =============================================================================
# CP / LP
# =============================================================================

def calcular_cp_lp(df):
    if 'SUPRIMENTO_INICIO' not in df.columns or 'SUPRIMENTO_TERMINO' not in df.columns:
        return df

    df['DIAS']  = (df['SUPRIMENTO_TERMINO'] - df['SUPRIMENTO_INICIO']).dt.days
    df['CP/LP'] = df['DIAS'].apply(classificar_cp_lp)

    return df

df = calcular_cp_lp(df)


# =============================================================================
# HORAS DO MÊS
# =============================================================================

def calcular_horas_mes(df):
    if 'MES' not in df.columns or 'SUPRIMENTO_INICIO' not in df.columns:
        return df

    df['MES'] = pd.to_numeric(df['MES'], errors='coerce')
    df['ANO'] = df['SUPRIMENTO_INICIO'].dt.year

    df['HORAS_MES'] = df.apply(
        lambda linha: total_horas_mes(linha['MES'], linha['ANO']),
        axis=1
    )

    return df

df = calcular_horas_mes(df)


# =============================================================================
# COLUNA FONTE
# =============================================================================

def tratar_fonte(df):
    if 'FONTE' not in df.columns:
        return df

    df['FONTE'] = df['FONTE'].replace({
        'Incentivada 50%':           'Incentivada-I5',
        'Cogeração Qualificada 50%': 'Incentivada-CQ5',
        'Incentivada 100%':          'Incentivada-I1',
        'Incentivada 0%':            'Incentivada-I0',
    })

    return df

df = tratar_fonte(df)


# =============================================================================
# COLUNA SUBMERCADO
# =============================================================================

def tratar_submercado(df):
    if 'SUBMERCADO' not in df.columns:
        return df

    df['SUBMERCADO'] = (
        df['SUBMERCADO']
        .astype(str)
        .str.strip()
        .str.upper()
        .replace({
            'N':     'NORTE',
            'S':     'SUL',
            'NE':    'NORDESTE',
            'SE/CO': 'SUDESTE',
        })
    )

    return df

df = tratar_submercado(df)


# =============================================================================
# COLUNA MODULAÇÃO
# =============================================================================

def tratar_modulacao(df):
    if 'MODULACAO WBC' not in df.columns:
        return df

    df['MODULACAO WBC'] = df['MODULACAO WBC'].replace({
        'C - Carga': 'CARGA',
        'F - Flat':  'FLAT',
        'DECLARADO': 'DECLARADA',
    })

    return df

df = tratar_modulacao(df)


# =============================================================================
# COLUNA MONTANTE MWh
# =============================================================================

def calcular_montante_mwh(df):
    if 'MONTANTE_MWH' not in df.columns:
        return df

    df['MONTANTE_MWH_NUM'] = pd.to_numeric(df['MONTANTE_MWH'], errors='coerce')

    df['MONTANTE MWh'] = df['MONTANTE_MWH_NUM'].apply(
        lambda valor: formatar_numero_br(valor, 3)
    )

    return df

df = calcular_montante_mwh(df)


# =============================================================================
# COLUNA MONTANTE MWm
# =============================================================================

def calcular_montante_mwm(df):
    if 'MONTANTE_MWH_NUM' not in df.columns or 'HORAS_MES' not in df.columns:
        return df

    df['MONTANTE_MWM_NUM'] = df['MONTANTE_MWH_NUM'] / df['HORAS_MES']

    df['MONTANTE MWm'] = df['MONTANTE_MWM_NUM'].apply(
        lambda valor: formatar_numero_br(valor, 6)
    )

    return df

df = calcular_montante_mwm(df)


# =============================================================================
# PREENCHE VAZIOS
# =============================================================================

df = df.fillna('-')


# =============================================================================
# CLIQ MÊS ANTERIOR
# =============================================================================

if arquivo_mes_anterior is not None:

    try:
        df_anterior = pd.read_excel(arquivo_mes_anterior)

        df_anterior.columns = [limpar_coluna(col) for col in df_anterior.columns]

        df_anterior = df_anterior.rename(columns={
            'CODIGO_WBC':  'BOLETA',
            'CODIGO_CCEE': 'Cliq Mês Anterior',
        })

        df['BOLETA']          = df['BOLETA'].astype(str).str.strip()
        df_anterior['BOLETA'] = df_anterior['BOLETA'].astype(str).str.strip()

        df = df.merge(
            df_anterior[['BOLETA', 'Cliq Mês Anterior']],
            on='BOLETA',
            how='left'
        )

        df['Cliq Mês Anterior'] = df['Cliq Mês Anterior'].fillna('-')

        st.success('✅ Cliq do mês anterior encontrado!')

    except Exception as erro:
        st.warning(f'⚠️ Erro ao buscar mês anterior: {erro}')


# =============================================================================
# EXIBIÇÃO
# =============================================================================

with st.expander('🔍 Ver colunas disponíveis'):
    st.write(df.columns.tolist())

colunas_existentes = [col for col in COLUNAS_EXIBICAO if col in df.columns]

st.subheader('Contratos Aprovados')
st.dataframe(df[colunas_existentes], hide_index=True, use_container_width=True)
