# =============================================================================
#
# ⚡ BOOK DE ENERGIA
# =============================================================================
#
# OBJETIVO
# --------
# Ler arquivos de contratos de energia, tratar os dados
# e exibir o resultado em uma interface Streamlit.
#
# ESTRUTURA DO PROJETO
# --------------------
# 1. IMPORTAÇÕES
# 2. CONFIGURAÇÕES
# 3. FUNÇÕES UTILITÁRIAS
# 4. REGRAS DE NEGÓCIO
# 5. PROCESSAMENTO
# 6. INTERFACE STREAMLIT
#
# =============================================================================


# =============================================================================
# 1. IMPORTAÇÕES
# =============================================================================

import streamlit as st
import pandas as pd
import unicodedata
import calendar


# =============================================================================
# 2. CONFIGURAÇÕES
# =============================================================================

# -----------------------------------------------------------------------------
# Arquivos permitidos no upload
# -----------------------------------------------------------------------------

TIPOS_ARQUIVO = ['xlsx', 'xlsm', 'csv']


# -----------------------------------------------------------------------------
# Colunas exibidas no resultado final
# -----------------------------------------------------------------------------

COLUNAS_EXIBICAO = [
    'BOLETA',
    'OPERACAO',
    'FONTE',
    'PARTE',
    'CONTRAPARTE',
    'CP/LP',
    'SUBMERCADO',
    'MONTANTE_MWH_FORMATADO',
    'MONTANTE_MWM_FORMATADO',
    'CLIQ PARADIGMA',
    'MODULACAO WBC',
    'MOD MIN',
    'MOD MAX',
]


# -----------------------------------------------------------------------------
# Renomeação de colunas
# nome_original -> nome_novo
# -----------------------------------------------------------------------------

MAPA_RENOMEACAO = {
    'PARTE_NOME_FANTASIA': 'PARTE',
    'MOVIMENTACAO': 'OPERACAO',
    'FONTE_CONTRATO': 'FONTE',
    'CODIGO_WBC': 'BOLETA',
    'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
    'QUANTATUALIZADA': 'MONTANTE_MWH',
    'CODIGO_CCEE': 'CLIQ PARADIGMA',
    'TIPO_DE_MODULACAO': 'MODULACAO WBC',
    'FLEXLIMITE_MODULACAOMAX': 'MOD MAX',
    'FLEXLIMITE_MODULACAOMIN': 'MOD MIN',
}


# -----------------------------------------------------------------------------
# Conversão de fontes
# -----------------------------------------------------------------------------

MAPA_FONTES = {
    'Incentivada 50%': 'Incentivada-I5',
    'Cogeração Qualificada 50%': 'Incentivada-CQ5',
    'Incentivada 100%': 'Incentivada-I1',
    'Incentivada 0%': 'Incentivada-I0',
}


# -----------------------------------------------------------------------------
# Conversão de submercados
# -----------------------------------------------------------------------------

MAPA_SUBMERCADOS = {
    'N': 'NORTE',
    'S': 'SUL',
    'NE': 'NORDESTE',
    'SE/CO': 'SUDESTE',
}


# -----------------------------------------------------------------------------
# Conversão de modulação
# -----------------------------------------------------------------------------

MAPA_MODULACAO = {
    'C - Carga': 'CARGA',
    'F - Flat': 'FLAT',
    'DECLARADO': 'DECLARADA',
}


# =============================================================================
# 3. FUNÇÕES UTILITÁRIAS
# =============================================================================

def limpar_coluna(texto):
    """
    Padroniza nomes de colunas:
    - remove espaços
    - transforma em maiúsculo
    - remove acentos
    """

    texto = str(texto).strip().upper()

    texto = (
        unicodedata
        .normalize('NFKD', texto)
        .encode('ASCII', 'ignore')
        .decode('utf-8')
    )

    return texto


def formatar_numero_br(valor, casas_decimais=2):
    """
    Formata números no padrão brasileiro.
    """

    try:
        return (
            f"{valor:,.{casas_decimais}f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except Exception:
        return '-'


def coluna_existe(df, coluna):
    """
    Verifica se uma coluna existe no DataFrame.
    """

    return coluna in df.columns


# =============================================================================
# 4. REGRAS DE NEGÓCIO
# =============================================================================

def tratar_fonte(valor):
    return MAPA_FONTES.get(valor, valor)


def tratar_submercado(valor):

    valor = str(valor).strip().upper()

    return MAPA_SUBMERCADOS.get(valor, valor)


def tratar_modulacao(valor):
    return MAPA_MODULACAO.get(valor, valor)


def calcular_cp_lp(dias):

    if pd.isna(dias):
        return '-'

    return 'LP' if dias > 31 else 'CP'


def calcular_horas_mes(mes, ano):

    try:
        dias_mes = calendar.monthrange(int(ano), int(mes))[1]
        return dias_mes * 24

    except Exception:
        return None


# =============================================================================
# 5. PROCESSAMENTO
# =============================================================================

def padronizar_colunas(df):
    """
    Padroniza nomes das colunas.
    """

    df.columns = [limpar_coluna(col) for col in df.columns]

    return df


def renomear_colunas(df):
    """
    Renomeia colunas para o padrão do book.
    """

    colunas_validas = {
        coluna_original: coluna_nova
        for coluna_original, coluna_nova in MAPA_RENOMEACAO.items()
        if coluna_original in df.columns
    }

    colunas_ausentes = (
        set(MAPA_RENOMEACAO.keys())
        - set(colunas_validas.keys())
    )

    if colunas_ausentes:
        st.warning(
            f"""
            ⚠️ Colunas não encontradas:

            {', '.join(sorted(colunas_ausentes))}
            """
        )

    return df.rename(columns=colunas_validas)


def aplicar_tratamentos(df):
    """
    Aplica tratamentos padronizados nas colunas.
    """

    tratamentos = {
        'FONTE': tratar_fonte,
        'SUBMERCADO': tratar_submercado,
        'MODULACAO WBC': tratar_modulacao,
    }

    for coluna, funcao in tratamentos.items():

        if coluna_existe(df, coluna):
            df[coluna] = df[coluna].apply(funcao)

    return df


def calcular_datas(df):
    """
    Calcula DIAS e classificação CP/LP.
    """

    colunas_data = (
        coluna_existe(df, 'SUPRIMENTO_INICIO')
        and
        coluna_existe(df, 'SUPRIMENTO_TERMINO')
    )

    if not colunas_data:
        return df

    df['SUPRIMENTO_INICIO'] = pd.to_datetime(
        df['SUPRIMENTO_INICIO'],
        errors='coerce'
    )

    df['SUPRIMENTO_TERMINO'] = pd.to_datetime(
        df['SUPRIMENTO_TERMINO'],
        errors='coerce'
    )

    df['DIAS'] = (
        df['SUPRIMENTO_TERMINO']
        - df['SUPRIMENTO_INICIO']
    ).dt.days

    df['CP/LP'] = df['DIAS'].apply(calcular_cp_lp)

    return df


def calcular_horas(df):
    """
    Calcula quantidade de horas do mês.
    """

    if not coluna_existe(df, 'MES'):
        return df

    df['MES'] = pd.to_numeric(df['MES'], errors='coerce')

    if coluna_existe(df, 'SUPRIMENTO_INICIO'):
        df['ANO'] = df['SUPRIMENTO_INICIO'].dt.year

    else:
        df['ANO'] = pd.Timestamp.today().year

    df['HORAS_MES'] = df.apply(
        lambda linha: calcular_horas_mes(
            linha['MES'],
            linha['ANO']
        ),
        axis=1
    )

    return df


def calcular_montantes(df):
    """
    Calcula MWh e MWm.
    """

    if not coluna_existe(df, 'MONTANTE_MWH'):
        return df

    # -------------------------------------------------------------------------
    # Conversão numérica
    # -------------------------------------------------------------------------

    df['MONTANTE_MWH_NUM'] = pd.to_numeric(
        df['MONTANTE_MWH'],
        errors='coerce'
    )

    # -------------------------------------------------------------------------
    # MWh formatado
    # -------------------------------------------------------------------------

    df['MONTANTE_MWH_FORMATADO'] = (
        df['MONTANTE_MWH_NUM']
        .apply(lambda valor: formatar_numero_br(valor, 3))
    )

    # -------------------------------------------------------------------------
    # MWm
    # -------------------------------------------------------------------------

    if coluna_existe(df, 'HORAS_MES'):

        df['MONTANTE_MWM_NUM'] = (
            df['MONTANTE_MWH_NUM']
            / df['HORAS_MES']
        )

        df['MONTANTE_MWM_FORMATADO'] = (
            df['MONTANTE_MWM_NUM']
            .apply(lambda valor: formatar_numero_br(valor, 6))
        )

    return df


def processar_contratos(df):
    """
    Pipeline principal de processamento.
    """

    df = df.copy()

    df = padronizar_colunas(df)

    df = renomear_colunas(df)

    df = aplicar_tratamentos(df)

    df = calcular_datas(df)

    df = calcular_horas(df)

    df = calcular_montantes(df)

    return df


# =============================================================================
# 6. INTERFACE STREAMLIT
# =============================================================================

st.set_page_config(
    page_title='Book de Energia',
    layout='wide'
)

st.title('⚡ Book de Energia')


# =============================================================================
# UPLOADS
# =============================================================================

coluna_esquerda, coluna_direita = st.columns(2)


# -----------------------------------------------------------------------------
# Upload principal
# -----------------------------------------------------------------------------

with coluna_esquerda:

    arquivo_aprovados = st.file_uploader(
        label='Contratos aprovados',
        type=TIPOS_ARQUIVO,
        key='aprovados'
    )


# -----------------------------------------------------------------------------
# Upload mês anterior
# -----------------------------------------------------------------------------

with coluna_direita:

    arquivo_mes_anterior = st.file_uploader(
        label='Contratos mês anterior',
        type=TIPOS_ARQUIVO,
        key='mes_anterior'
    )


# =============================================================================
# VALIDAÇÃO
# =============================================================================

if arquivo_aprovados is None:

    st.info(
        '📂 Faça upload do arquivo de contratos aprovados.'
    )

    st.stop()


# =============================================================================
# LEITURA DO ARQUIVO PRINCIPAL
# =============================================================================

try:

    df_aprovados = pd.read_excel(
        arquivo_aprovados,
        skiprows=8
    )

    st.success(
        '✅ Arquivo de contratos aprovados carregado!'
    )

except Exception as erro:

    st.error(
        f'❌ Erro ao ler arquivo: {erro}'
    )

    st.stop()


# =============================================================================
# LEITURA DO MÊS ANTERIOR
# =============================================================================

df_mes_anterior = None

if arquivo_mes_anterior is not None:

    try:

        df_mes_anterior = pd.read_excel(
            arquivo_mes_anterior
        )

        st.success(
            '✅ Arquivo do mês anterior carregado!'
        )

    except Exception as erro:

        st.warning(
            f'⚠️ Erro ao ler mês anterior: {erro}'
        )


# =============================================================================
# PROCESSAMENTO
# =============================================================================

try:

    df_processado = processar_contratos(df_aprovados)

except Exception as erro:

    st.error(
        f'❌ Erro no processamento: {erro}'
    )

    st.stop()


# =============================================================================
# DEBUG
# =============================================================================

with st.expander('🔍 Ver colunas encontradas'):

    st.write(
        df_processado.columns.tolist()
    )


# =============================================================================
# EXIBIÇÃO — CONTRATOS APROVADOS
# =============================================================================

colunas_existentes = [
    coluna
    for coluna in COLUNAS_EXIBICAO
    if coluna in df_processado.columns
]


colunas_faltando = (
    set(COLUNAS_EXIBICAO)
    - set(colunas_existentes)
)

if colunas_faltando:

    st.warning(
        f"""
        ⚠️ Colunas não encontradas:

        {', '.join(sorted(colunas_faltando))}
        """
    )


st.subheader('Contratos Aprovados')

st.dataframe(
    df_processado[colunas_existentes],
    hide_index=True,
    use_container_width=True
)


# =============================================================================
# EXIBIÇÃO — MÊS ANTERIOR
# =============================================================================

if df_mes_anterior is not None:

    st.subheader('Contratos Mês Anterior')

    st.dataframe(
        df_mes_anterior,
        hide_index=True,
        use_container_width=True
    )
