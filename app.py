# =============================================================================
#
# ⚡ BOOK DE ENERGIA
#
# =============================================================================

import streamlit as st
import pandas as pd
import unicodedata
import calendar
import zipfile
import io


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
    'MODULACAO WBC',
    'MOD MIN',
    'MOD MAX',
    'Cliq Mês Anterior',
    'VENDEDOR',
    'COMPRADOR',
    'CLIQ CCEE'
]


# =============================================================================
# FUNÇÕES DE TEXTO
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


def formatar_cnpj(valor):

    try:

        digits = ''.join(filter(str.isdigit, str(valor)))

        digits = digits.zfill(14)

        return (
            f"{digits[:2]}."
            f"{digits[2:5]}."
            f"{digits[5:8]}/"
            f"{digits[8:12]}-"
            f"{digits[12:]}"
        )

    except Exception:

        return str(valor)


# =============================================================================
# FUNÇÕES DE NEGÓCIO
# =============================================================================

def classificar_cp_lp(dias):

    if pd.isna(dias):
        return '-'

    return 'LP' if dias > 31 else 'CP'


def total_horas_mes(mes, ano):

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


# =============================================================================
# UPLOADS
# =============================================================================

st.set_page_config(
    page_title='Book de Energia',
    layout='wide'
)

st.title('⚡ Book de Energia')

col1, col2, col3 = st.columns(3)

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

with col3:

    arquivo_zip = st.file_uploader(
        label='ZIP CLIQ BISMUT',
        type=['zip'],
        key='zip_bismut'
    )

if arquivo_aprovados is None:

    st.info('📂 Faça upload do arquivo principal.')

    st.stop()


# =============================================================================
# LEITURA ARQUIVO PRINCIPAL
# =============================================================================

try:

    df = pd.read_excel(
        arquivo_aprovados,
        skiprows=8
    )

    st.success('✅ Arquivo principal carregado!')

except Exception as erro:

    st.error(f'❌ Erro ao ler arquivo principal: {erro}')

    st.stop()


# =============================================================================
# LEITURA ZIP BISMUT
# =============================================================================

df_bismut = None

if arquivo_zip is not None:

    try:

        with zipfile.ZipFile(arquivo_zip) as zip_ref:

            arquivos_zip = zip_ref.namelist()

            arquivo_bismut = next(
                (
                    arquivo
                    for arquivo in arquivos_zip
                    if 'cceal_firme_' in arquivo.lower()
                ),
                None
            )

            if arquivo_bismut is None:

                st.warning(
                    '⚠️ Arquivo cceal_firme não encontrado.'
                )

            else:

                with zip_ref.open(arquivo_bismut) as arquivo:

                    if arquivo_bismut.lower().endswith('.csv'):

                        try:

                            df_bismut = pd.read_csv(
                                arquivo,
                                encoding='utf-8',
                                sep=';'
                            )

                        except:

                            arquivo.seek(0)

                            df_bismut = pd.read_csv(
                                arquivo,
                                encoding='latin1',
                                sep=';'
                            )

                    elif arquivo_bismut.lower().endswith(
                        ('.xlsx', '.xlsm', '.xls')
                    ):

                        df_bismut = pd.read_excel(arquivo)

                st.success(
                    f'✅ Arquivo encontrado: {arquivo_bismut}'
                )

    except Exception as erro:

        st.error(f'❌ Erro ao ler ZIP: {erro}')


# =============================================================================
# LIMPEZA NOMES COLUNAS
# =============================================================================

def limpar_nomes_colunas(df):

    df.columns = [
        limpar_coluna(col)
        for col in df.columns
    ]

    return df


df = limpar_nomes_colunas(df)


# =============================================================================
# RENOMEAÇÃO COLUNAS
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
        'SIGLA_CCEE_VENDEDOR':       'VENDEDOR',
        'SIGLA_CCEE_COMPRADOR':      'COMPRADOR'

    })

    return df


df = renomear_colunas(df)

df = tratar_cnpj(df)


# =============================================================================
# DATAS
# =============================================================================

def converter_datas(df):

    if (
        'SUPRIMENTO_INICIO' not in df.columns
        or
        'SUPRIMENTO_TERMINO' not in df.columns
    ):
        return df

    df['SUPRIMENTO_INICIO'] = pd.to_datetime(
        df['SUPRIMENTO_INICIO'],
        errors='coerce'
    )

    df['SUPRIMENTO_TERMINO'] = pd.to_datetime(
        df['SUPRIMENTO_TERMINO'],
        errors='coerce'
    )

    return df


df = converter_datas(df)


# =============================================================================
# CP / LP
# =============================================================================

def calcular_cp_lp(df):

    if (
        'SUPRIMENTO_INICIO' not in df.columns
        or
        'SUPRIMENTO_TERMINO' not in df.columns
    ):
        return df

    df['DIAS'] = (
        df['SUPRIMENTO_TERMINO']
        -
        df['SUPRIMENTO_INICIO']
    ).dt.days

    df['CP/LP'] = (
        df['DIAS']
        .apply(classificar_cp_lp)
    )

    return df


df = calcular_cp_lp(df)


# =============================================================================
# HORAS MÊS
# =============================================================================

def calcular_horas_mes(df):

    if (
        'MES' not in df.columns
        or
        'SUPRIMENTO_INICIO' not in df.columns
    ):
        return df

    df['MES'] = pd.to_numeric(
        df['MES'],
        errors='coerce'
    )

    df['ANO'] = df['SUPRIMENTO_INICIO'].dt.year

    df['HORAS_MES'] = df.apply(
        lambda linha: total_horas_mes(
            linha['MES'],
            linha['ANO']
        ),
        axis=1
    )

    return df


df = calcular_horas_mes(df)


# =============================================================================
# FONTE
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
# SUBMERCADO
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
# MODULAÇÃO
# =============================================================================

def tratar_modulacao(df):

    if 'MODULACAO WBC' not in df.columns:
        return df

    df['MODULACAO WBC'] = (
        df['MODULACAO WBC']
        .replace({

            'C - Carga': 'CARGA',
            'F - Flat':  'FLAT',
            'DECLARADO': 'DECLARADA',

        })
    )

    return df


df = tratar_modulacao(df)


# =============================================================================
# MONTANTE MWh
# =============================================================================

def calcular_montante_mwh(df):

    if 'MONTANTE_MWH' not in df.columns:
        return df

    df['MONTANTE_MWH_NUM'] = pd.to_numeric(
        df['MONTANTE_MWH'],
        errors='coerce'
    )

    df['MONTANTE MWh'] = (
        df['MONTANTE_MWH_NUM']
        .apply(lambda valor: formatar_numero_br(valor, 3))
    )

    return df


df = calcular_montante_mwh(df)


# =============================================================================
# MONTANTE MWm
# =============================================================================

def calcular_montante_mwm(df):

    if (
        'MONTANTE_MWH_NUM' not in df.columns
        or
        'HORAS_MES' not in df.columns
    ):
        return df

    df['MONTANTE_MWM_NUM'] = (
        df['MONTANTE_MWH_NUM']
        /
        df['HORAS_MES']
    )

    df['MONTANTE MWm'] = (
        df['MONTANTE_MWM_NUM']
        .apply(lambda valor: formatar_numero_br(valor, 6))
    )

    return df


df = calcular_montante_mwm(df)


# =============================================================================
# CLIQ MÊS ANTERIOR
# =============================================================================

if arquivo_mes_anterior is not None:

    try:

        df_anterior = pd.read_excel(
            arquivo_mes_anterior
        )

        df_anterior.columns = [
            limpar_coluna(col)
            for col in df_anterior.columns
        ]

        df_anterior = df_anterior.rename(columns={

            'CODIGO_WBC':  'BOLETA',
            'CODIGO_CCEE': 'Cliq Mês Anterior',

        })

        df['BOLETA'] = (
            df['BOLETA']
            .astype(str)
            .str.strip()
        )

        df_anterior['BOLETA'] = (
            df_anterior['BOLETA']
            .astype(str)
            .str.strip()
        )

        df = df.merge(
            df_anterior[
                ['BOLETA', 'Cliq Mês Anterior']
            ],
            on='BOLETA',
            how='left'
        )

        df['Cliq Mês Anterior'] = (
            df['Cliq Mês Anterior']
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
# GARANTIR CLIQ MÊS ANTERIOR (caso arquivo não carregado)
# =============================================================================

if 'Cliq Mês Anterior' not in df.columns:
    df['Cliq Mês Anterior'] = '-'


# =============================================================================
# PREENCHER VAZIOS
# =============================================================================

df = df.fillna('-')


# =============================================================================
# MATCH CLIQ BISMUT
# =============================================================================

# Garantir coluna CLIQ CCEE sempre existe
df['CLIQ CCEE'] = '-'

if df_bismut is not None:

    try:

        df_bismut.columns = [
            limpar_coluna(col)
            for col in df_bismut.columns
        ]

        # Verificar se coluna PARTE existe no bismut
        if 'PARTE' not in df_bismut.columns:
            st.warning('⚠️ Coluna PARTE não encontrada no arquivo Bismut.')

        else:

            df_bismut_filtrado = df_bismut[
                df_bismut['PARTE']
                .astype(str)
                .str.strip()
                .str.upper()
                .str.contains('BISMUT', na=False)
            ].copy()

            if df_bismut_filtrado.empty:
                st.warning('⚠️ Nenhuma linha da Bismut encontrada no arquivo.')

            else:

                df['CLIQ PARADIGMA'] = (
                    df['CLIQ PARADIGMA']
                    .astype(str)
                    .str.strip()
                )

                df['Cliq Mês Anterior'] = (
                    df['Cliq Mês Anterior']
                    .astype(str)
                    .str.strip()
                )

                df['VENDEDOR'] = (
                    df['VENDEDOR']
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df['COMPRADOR'] = (
                    df['COMPRADOR']
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df_bismut_filtrado['CODIGO_CONTRATO'] = (
                    df_bismut_filtrado['CODIGO_CONTRATO']
                    .astype(str)
                    .str.strip()
                )

                df_bismut_filtrado['SIGLA_PERFIL_VENDEDOR'] = (
                    df_bismut_filtrado['SIGLA_PERFIL_VENDEDOR']
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                df_bismut_filtrado['SIGLA_PERFIL_COMPRADOR'] = (
                    df_bismut_filtrado['SIGLA_PERFIL_COMPRADOR']
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )

                def localizar_cliq(linha):

                    cliq_atual    = linha['CLIQ PARADIGMA']
                    cliq_anterior = linha['Cliq Mês Anterior']
                    vendedor      = linha['VENDEDOR']
                    comprador     = linha['COMPRADOR']

                    valores_vazios = {'', '-', 'nan', 'none', 'NAN', 'NONE'}

                    # Tenta com CLIQ PARADIGMA atual
                    if cliq_atual.lower() not in {'', '-', 'nan', 'none'}:

                        resultado = df_bismut_filtrado[
                            (df_bismut_filtrado['CODIGO_CONTRATO']        == cliq_atual)
                            &
                            (df_bismut_filtrado['SIGLA_PERFIL_VENDEDOR']  == vendedor)
                            &
                            (df_bismut_filtrado['SIGLA_PERFIL_COMPRADOR'] == comprador)
                        ]

                        if not resultado.empty:
                            return cliq_atual

                    # Tenta com CLIQ do mês anterior
                    if cliq_anterior.lower() not in {'', '-', 'nan', 'none'}:

                        resultado = df_bismut_filtrado[
                            (df_bismut_filtrado['CODIGO_CONTRATO']        == cliq_anterior)
                            &
                            (df_bismut_filtrado['SIGLA_PERFIL_VENDEDOR']  == vendedor)
                            &
                            (df_bismut_filtrado['SIGLA_PERFIL_COMPRADOR'] == comprador)
                        ]

                        if not resultado.empty:
                            return cliq_anterior

                    return 'VERIFICAR'

                df['CLIQ CCEE'] = df.apply(localizar_cliq, axis=1)

                total         = len(df)
                encontrados   = (df['CLIQ CCEE'] != 'VERIFICAR').sum()
                verificar     = (df['CLIQ CCEE'] == 'VERIFICAR').sum()

                st.success(
                    f'✅ Match Bismut realizado! '
                    f'{encontrados}/{total} encontrados | '
                    f'{verificar} para VERIFICAR'
                )

    except Exception as erro:

        st.warning(f'⚠️ Erro no match Bismut: {erro}')


# =============================================================================
# EXIBIÇÃO
# =============================================================================

with st.expander('🔍 Ver colunas disponíveis'):

    st.write(df.columns.tolist())

colunas_existentes = [
    col
    for col in COLUNAS_EXIBICAO
    if col in df.columns
]

st.subheader('Contratos Aprovados')

st.dataframe(
    df[colunas_existentes],
    hide_index=True,
    use_container_width=True
)
