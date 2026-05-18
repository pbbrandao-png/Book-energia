import streamlit as st
import pandas as pd
import unicodedata


# =========================================================
# FUNÇÕES DE PADRONIZAÇÃO
# =========================================================

def limpar_coluna(texto):

    texto = str(texto).strip().upper()

    texto = (
        unicodedata
        .normalize('NFKD', texto)
        .encode('ASCII', 'ignore')
        .decode('utf-8')
           )

    return texto


# =========================================================
# FUNÇÕES DE TRATAMENTO
# =========================================================

# ─────────────────────────────────────────────────────────
# FONTE
# ─────────────────────────────────────────────────────────

def tratar_fonte(valor):

    mapa = {
        'Incentivada 50%': 'Incentivada-I5',
        'Cogeração Qualificada 50%': 'Incentivada-CQ5',
        'Incentivada 100%': 'Incentivada-I1',
        'Incentivada 0%': 'Incentivada-I0',
    }

    return mapa.get(valor, valor)


# ─────────────────────────────────────────────────────────
# SUBMERCADO
# ─────────────────────────────────────────────────────────

def tratar_submercado(valor):

    valor = str(valor).strip().upper()

    mapa = {
        'N': 'NORTE',
        'S': 'SUL',
        'NE': 'NORDESTE',
        'SE/CO': 'SUDESTE',
    }

    return mapa.get(valor, valor)

# ─────────────────────────────────────────────────────────
# CP / LP
# ─────────────────────────────────────────────────────────

def calcular_cp_lp(dias):

    if dias > 31:
        return 'LP'

    return 'CP'


# ─────────────────────────────────────────────────────────
# FORMATA 3 CASAS
# ─────────────────────────────────────────────────────────

def formatar_3_casas(valor):

    try:

        return (
            f"{valor:,.3f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except:

        return valor


# ─────────────────────────────────────────────────────────
# FORMATA 6 CASAS
# ─────────────────────────────────────────────────────────

def formatar_6_casas(valor):

    try:

        return (
            f"{valor:,.6f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except:

        return valor


# ─────────────────────────────────────────────────────────
# HORAS DO MÊS
# ─────────────────────────────────────────────────────────

def horas_mes(mes):

    mapa = {
        1: 744,
        2: 672,
        3: 744,
        4: 720,
        5: 744,
        6: 720,
        7: 744,
        8: 744,
        9: 720,
        10: 744,
        11: 720,
        12: 744
    }

    return mapa.get(mes)

# ─────────────────────────────────────────────────────────
# MODULACAO
# ─────────────────────────────────────────────────────────

def tratar_modulacao(valor):

    mapa = {
        'C - Carga': 'CARGA',
        'F - Flat': 'FLAT',
        'DECLARADO': 'DECLARADA',
            }

    return mapa.get(valor, valor)




# =========================================================
# TÍTULO
# =========================================================

st.title("Livro de Energia - Abril/2026")


# =========================================================
# UPLOAD
# =========================================================

arquivo = st.file_uploader(
    "Contratos aprovados",
    type=['xlsx', 'csv', 'xlsm'],

    arquivo = st.file_uploader(
    "Contratos mês anterior",
    type=['xlsx', 'csv', 'xlsm']
)


# =========================================================
# PROCESSAMENTO
# =========================================================

if arquivo is not None:

    st.success("Arquivo carregado com sucesso!")

    # =====================================================
    # LEITURA
    # =====================================================

    df_contratos_aprovados = pd.read_excel(
        arquivo,
        skiprows=8
    )
    df_contratos_aprovados = df_contratos_aprovados.fillna("-")
    # =====================================================
    # PADRONIZAÇÃO DE COLUNAS
    # =====================================================

    df_contratos_aprovados.columns = [limpar_coluna(col) for col in df_contratos_aprovados.columns]

    # =====================================================
    # DEBUG
    # =====================================================

    st.write("Colunas encontradas:")
    st.write(df_contratos_aprovados.columns.tolist())

    # =====================================================
    # RENOMEAÇÃO DE COLUNAS
    # =====================================================

    df_contratos_aprovados = df_contratos_aprovados.rename(
        columns={
            'PARTE_NOME_FANTASIA': 'PARTE',
            'MOVIMENTACAO': 'OPERACAO',
            'FONTE_CONTRATO': 'FONTE',
            'CODIGO_WBC': 'BOLETA',
            'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
            'QUANTATUALIZADA': 'MONTANTE_MWH',
            'CODIGO_CCEE': 'CLIQ PARADIGMA',
            'TIPO_DE_MODULACAO':'MODULACAO WBC',
            'FLEXLIMITE_MODULACAOMAX':'MOD MAX',
            'FLEXLIMITE_MODULACAOMIN':'MOD MIN'
        }
    )

    # =====================================================
    # TRATAMENTO FONTE
    # =====================================================

    if 'FONTE' in df_contratos_aprovados.columns:

        df_contratos_aprovados['FONTE'] = df_contratos_aprovados['FONTE'].apply(
            tratar_fonte
        )

    # =====================================================
    # TRATAMENTO SUBMERCADO
    # =====================================================

    if 'SUBMERCADO' in df_contratos_aprovados.columns:

        df_contratos_aprovados['SUBMERCADO'] = df_contratos_aprovados['SUBMERCADO'].apply(
            tratar_submercado
        )
  
    # =====================================================
    # TRATAMENTO DATAS
    # =====================================================

    if (
        'SUPRIMENTO_INICIO' in df_contratos_aprovados.columns
        and
        'SUPRIMENTO_TERMINO' in df_contratos_aprovados.columns
    ):

        df_contratos_aprovados['SUPRIMENTO_INICIO'] = pd.to_datetime(
            df_contratos_aprovados['SUPRIMENTO_INICIO'],
            errors='coerce'
        )

        df_contratos_aprovados['SUPRIMENTO_TERMINO'] = pd.to_datetime(
            df_contratos_aprovados['SUPRIMENTO_TERMINO'],
            errors='coerce'
        )

        # CALCULA DIAS

        df_contratos_aprovados['DIAS'] = (
            df_contratos_aprovados['SUPRIMENTO_TERMINO']
            - df_contratos_aprovados['SUPRIMENTO_INICIO']
        ).dt.days

        # CALCULA CP / LP

        df_contratos_aprovados['CP/LP'] = df_contratos_aprovados['DIAS'].apply(
            calcular_cp_lp
        )

    # =====================================================
    # HORAS DO MÊS
    # =====================================================

    if 'MES' in df_contratos_aprovados.columns:

        df_contratos_aprovados['MES'] = pd.to_numeric(
            df_contratos_aprovados['MES'],
            errors='coerce'
        )

        df_contratos_aprovados['HORAS_MES'] = df_contratos_aprovados['MES'].apply(
            horas_mes
        )

    # =====================================================
    # TRATAMENTO MONTANTE
    # =====================================================

    if 'MONTANTE_MWH' in df_contratos_aprovados.columns:

        # GARANTE NÚMERO

        df_contratos_aprovados['MONTANTE_MWH'] = pd.to_numeric(
            df_contratos_aprovados['MONTANTE_MWH'],
            errors='coerce'
        )

        # CALCULA MWm

        if 'HORAS_MES' in df_contratos_aprovados.columns:

            df_contratos_aprovados['MONTANTE_MWM'] = (
                df_contratos_aprovados['MONTANTE_MWH']
                / df_contratos_aprovados['HORAS_MES']
            )

            # FORMATA MWm (6 CASAS)

            df_contratos_aprovados['MONTANTE_MWM'] = df_contratos_aprovados[
                'MONTANTE_MWM'
            ].apply(formatar_6_casas)

        # FORMATA MWh (3 CASAS)

        df_contratos_aprovados['MONTANTE_MWH'] = df_contratos_aprovados[
            'MONTANTE_MWH'
        ].apply(formatar_3_casas)
    # =====================================================
    # TRATAMENTO MODULACAO
    # =====================================================

    if 'MODULACAO WBC' in df_contratos_aprovados.columns:

        df_contratos_aprovados['MODULACAO WBC'] = df_contratos_aprovados['MODULACAO WBC'].apply(
            tratar_modulacao
        )

    # =====================================================
    # RENOMEAÇÃO VISUAL
    # =====================================================

    df_contratos_aprovados = df_contratos_aprovados.rename(
        columns={
            'MONTANTE_MWH': 'MONTANTE MWh',
            'MONTANTE_MWM': 'MONTANTE MWm'
        }
    )

    # =====================================================
    # COLUNAS EXIBIDAS
    # =====================================================

    colunas_desejadas = [
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
        'MODULACAO WBC',
        'MOD MIN',
        'MOD MAX'
    ]

    # =====================================================
    # VALIDAÇÃO DE COLUNAS
    # =====================================================

    colunas_existentes = [
        col
        for col in colunas_desejadas
        if col in df_contratos_aprovados.columns
    ]

    # =====================================================
    # EXIBIÇÃO
    # =====================================================

    st.dataframe(
        df_contratos_aprovados[colunas_existentes],
        hide_index=True
    )
