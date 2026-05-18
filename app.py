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
    'SUBMERCADO',
    'MONTANTE MWh',
    'MONTANTE MWm',
    'CLIQ PARADIGMA',
    'Cliq Mês Anterior',
    'MODULACAO WBC',
    'MOD MIN',
    'MOD MAX',
    'Cliq Mês Anterior'
    
]


# =============================================================================
# FUNÇÕES UTILITÁRIAS
# =============================================================================

def limpar_coluna(texto):

    texto = str(texto).strip().upper()

    texto = (
        unicodedata
        .normalize('NFKD', texto)
        .encode('ASCII', 'ignore')
        .decode('utf-8')
    )

    return texto


def formatar_numero_br(valor, casas=2):

    try:

        return (
            f"{valor:,.{casas}f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except Exception:

        return '-'


def calcular_cp_lp(dias):

    if pd.isna(dias):
        return '-'

    return 'LP' if dias > 31 else 'CP'


def calcular_horas_mes(mes, ano):

    try:

        dias_mes = calendar.monthrange(
            int(ano),
            int(mes)
        )[1]

        return dias_mes * 24

    except Exception:

        return None


# =============================================================================
# PROCESSAMENTO
# =============================================================================

def processar_contratos(df):

    # =========================================================================
    # CÓPIA
    # =========================================================================

    df = df.copy()


    # =========================================================================
    # PADRONIZAÇÃO DAS COLUNAS
    # =========================================================================

    df.columns = [
        limpar_coluna(col)
        for col in df.columns
    ]


    # =========================================================================
    # RENOMEAÇÃO
    # =========================================================================

    df = df.rename(columns={

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

    })


    # =========================================================================
    # FONTE
    # =========================================================================

    if 'FONTE' in df.columns:

        df['FONTE'] = (

            df['FONTE']

            .replace({

                'Incentivada 50%': 'Incentivada-I5',
                'Cogeração Qualificada 50%': 'Incentivada-CQ5',
                'Incentivada 100%': 'Incentivada-I1',
                'Incentivada 0%': 'Incentivada-I0',

            })
        )


    # =========================================================================
    # SUBMERCADO
    # =========================================================================

    if 'SUBMERCADO' in df.columns:

        df['SUBMERCADO'] = (

            df['SUBMERCADO']

            .astype(str)

            .str.strip()

            .str.upper()

            .replace({

                'N': 'NORTE',
                'S': 'SUL',
                'NE': 'NORDESTE',
                'SE/CO': 'SUDESTE',

            })
        )


    # =========================================================================
    # MODULAÇÃO
    # =========================================================================

    if 'MODULACAO WBC' in df.columns:

        df['MODULACAO WBC'] = (

            df['MODULACAO WBC']

            .replace({

                'C - Carga': 'CARGA',
                'F - Flat': 'FLAT',
                'DECLARADO': 'DECLARADA',

            })
        )


    # =========================================================================
    # DATAS
    # =========================================================================

    if (
        'SUPRIMENTO_INICIO' in df.columns
        and
        'SUPRIMENTO_TERMINO' in df.columns
    ):

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
            -
            df['SUPRIMENTO_INICIO']

        ).dt.days

        df['CP/LP'] = (
            df['DIAS']
            .apply(calcular_cp_lp)
        )


    # =========================================================================
    # HORAS DO MÊS
    # =========================================================================

    if 'MES' in df.columns:

        df['MES'] = pd.to_numeric(
            df['MES'],
            errors='coerce'
        )

        if 'SUPRIMENTO_INICIO' in df.columns:

            df['ANO'] = (
                df['SUPRIMENTO_INICIO']
                .dt.year
            )

        else:

            df['ANO'] = (
                pd.Timestamp.today().year
            )

        df['HORAS_MES'] = df.apply(

            lambda linha:

            calcular_horas_mes(
                linha['MES'],
                linha['ANO']
            ),

            axis=1
        )


    # =========================================================================
    # MONTANTE
    # =========================================================================

    if 'MONTANTE_MWH' in df.columns:

        df['MONTANTE_MWH_NUM'] = pd.to_numeric(
            df['MONTANTE_MWH'],
            errors='coerce'
        )

        df['MONTANTE MWh'] = (

            df['MONTANTE_MWH_NUM']

            .apply(
                lambda valor:
                formatar_numero_br(valor, 3)
            )
        )

        if 'HORAS_MES' in df.columns:

            df['MONTANTE_MWM_NUM'] = (

                df['MONTANTE_MWH_NUM']
                /
                df['HORAS_MES']

            )

            df['MONTANTE MWm'] = (

                df['MONTANTE_MWM_NUM']

                .apply(
                    lambda valor:
                    formatar_numero_br(valor, 6)
                )
            )


    # =========================================================================
    # VAZIOS
    # =========================================================================

    df = df.fillna('-')


    return df


# =============================================================================
# STREAMLIT
# =============================================================================

st.set_page_config(

    page_title='Book de Energia',

    layout='wide'

)

st.title('⚡ Book de Energia')


# =============================================================================
# UPLOADS
# =============================================================================

col1, col2 = st.columns(2)


# -----------------------------------------------------------------------------
# CONTRATOS APROVADOS
# -----------------------------------------------------------------------------

with col1:

    arquivo_aprovados = st.file_uploader(

        label='Contratos aprovados',

        type=TIPOS_ARQUIVO,

        key='aprovados'
    )


# -----------------------------------------------------------------------------
# MÊS ANTERIOR
# -----------------------------------------------------------------------------

with col2:

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
        '📂 Faça upload do arquivo principal.'
    )

    st.stop()


# =============================================================================
# LEITURA PRINCIPAL
# =============================================================================

try:

    df_aprovados = pd.read_excel(

        arquivo_aprovados,

        skiprows=8

    )

    st.success(
        '✅ Arquivo principal carregado!'
    )

except Exception as erro:

    st.error(
        f'❌ Erro ao ler arquivo principal: {erro}'
    )

    st.stop()


# =============================================================================
# PROCESSAMENTO PRINCIPAL
# =============================================================================

try:

    df_processado = processar_contratos(
        df_aprovados
    )

except Exception as erro:

    st.error(
        f'❌ Erro no processamento: {erro}'
    )

    st.stop()


# =============================================================================
# PROCX — MÊS ANTERIOR
# =============================================================================

if arquivo_mes_anterior is not None:

    try:

        # ---------------------------------------------------------------------
        # LEITURA
        # ---------------------------------------------------------------------

        df_mes_anterior = pd.read_excel(
            arquivo_mes_anterior
        )


        # ---------------------------------------------------------------------
        # PADRONIZAÇÃO DAS COLUNAS
        # ---------------------------------------------------------------------

        df_mes_anterior.columns = [

            limpar_coluna(col)

            for col in df_mes_anterior.columns

        ]


        # ---------------------------------------------------------------------
        # RENOMEAÇÃO
        # CODIGO_WBC -> BOLETA
        # CODIGO_CCEE -> Cliq Mês Anterior
        # ---------------------------------------------------------------------

        df_mes_anterior = df_mes_anterior.rename(columns={

            'CODIGO_WBC': 'BOLETA',
            'CODIGO_CCEE': 'Cliq Mês Anterior',

        })


        # ---------------------------------------------------------------------
        # PADRONIZAÇÃO DA BOLETA
        # ---------------------------------------------------------------------

        df_processado['BOLETA'] = (

            df_processado['BOLETA']

            .astype(str)

            .str.strip()

        )

        df_mes_anterior['BOLETA'] = (

            df_mes_anterior['BOLETA']

            .astype(str)

            .str.strip()

        )


        # ---------------------------------------------------------------------
        # PROCX / MERGE
        # ---------------------------------------------------------------------

        df_processado = df_processado.merge(

            df_mes_anterior[[
                'BOLETA',
                'Cliq Mês Anterior'
            ]],

            on='BOLETA',

            how='left'

        )


        # ---------------------------------------------------------------------
        # SUBSTITUI VAZIOS
        # ---------------------------------------------------------------------

        df_processado['Cliq Mês Anterior'] = (

            df_processado['Cliq Mês Anterior']

            .fillna('-')

        )

        st.success(
            '✅ Cliq do mês anterior encontrado!'
        )

    except Exception as erro:

        st.warning(
            f'⚠️ Erro ao buscar mês anterior: {erro}'
        )


# =============================================================================
# DEBUG
# =============================================================================

with st.expander('🔍 Ver colunas disponíveis'):

    st.write(
        df_processado.columns.tolist()
    )


# =============================================================================
# COLUNAS EXISTENTES
# =============================================================================

colunas_existentes = [

    coluna

    for coluna in COLUNAS_EXIBICAO

    if coluna in df_processado.columns

]


# =============================================================================
# EXIBIÇÃO FINAL
# =============================================================================

st.subheader('Contratos Aprovados')

st.dataframe(

    df_processado[colunas_existentes],

    hide_index=True,

    use_container_width=True

)
