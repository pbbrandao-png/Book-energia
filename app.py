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
# FORMATA MONTANTE
# ─────────────────────────────────────────────────────────

def formatar_mwh(valor):

    try:

        return (
            f"{valor:,.3f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except:

        return valor


# =========================================================
# TÍTULO
# =========================================================

st.title("Livro de Energia - Abril/2026")


# =========================================================
# UPLOAD
# =========================================================

arquivo = st.file_uploader(
    "Contratos aprovados",
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

    df = pd.read_excel(
        arquivo,
        skiprows=8
    )

    # =====================================================
    # PADRONIZAÇÃO DE COLUNAS
    # =====================================================

    df.columns = [limpar_coluna(col) for col in df.columns]

    # =====================================================
    # DEBUG
    # =====================================================

    st.write("Colunas encontradas:")
    st.write(df.columns.tolist())

    # =====================================================
    # RENOMEAÇÃO DE COLUNAS
    # =====================================================

    df = df.rename(
        columns={
            'PARTE_NOME_FANTASIA': 'PARTE',
            'MOVIMENTACAO': 'OPERACAO',
            'FONTE_CONTRATO': 'FONTE',
            'CODIGO_WBC': 'BOLETA',
            'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
            'QUANTATUALIZADA': 'MONTANTE_MWH'
        }
    )

    # =====================================================
    # TRATAMENTO FONTE
    # =====================================================

    if 'FONTE' in df.columns:

        df['FONTE'] = df['FONTE'].apply(
            tratar_fonte
        )

    # =====================================================
    # TRATAMENTO SUBMERCADO
    # =====================================================

    if 'SUBMERCADO' in df.columns:

        df['SUBMERCADO'] = df['SUBMERCADO'].apply(
            tratar_submercado
        )

    # =====================================================
    # TRATAMENTO DATAS
    # =====================================================

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

        # CALCULA DIAS

        df['DIAS'] = (
            df['SUPRIMENTO_TERMINO']
            - df['SUPRIMENTO_INICIO']
        ).dt.days

        # CALCULA CP / LP

        df['CP/LP'] = df['DIAS'].apply(
            calcular_cp_lp
        )

    # =====================================================
    # FORMATAÇÃO MONTANTE
    # =====================================================

    if 'MONTANTE_MWH' in df.columns:

        df['MONTANTE_MWH'] = df['MONTANTE_MWH'].apply(
            formatar_mwh
        )

    # =====================================================
    # RENOMEAÇÃO VISUAL
    # =====================================================

    df = df.rename(
        columns={
            'MONTANTE_MWH': 'MONTANTE MWh'
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
        'MONTANTE MWh'
    ]

    # =====================================================
    # VALIDAÇÃO DE COLUNAS
    # =====================================================

    colunas_existentes = [
        col
        for col in colunas_desejadas
        if col in df.columns
    ]

    # =====================================================
    # EXIBIÇÃO
    # =====================================================

    st.dataframe(
        df[colunas_existentes],
        hide_index=True
    )
