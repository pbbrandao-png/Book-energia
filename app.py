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

# CLIQs que pertencem à Matrix — buscam no ccear_q
CLIQS_MATRIX = {
    '2813298', '2813299', '2813300', '2813301', '2813302',
    '2813303', '2813304', '2813305', '4159778', '4159779',
    '4159780', '4686267', '4686268', '4686269', '4686270'
}

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


def normalizar_codigo(valor):
    """Converte código para string sem decimais (ex: 1255694.0 → '1255694')."""

    s = str(valor).strip()

    if s.replace('.', '').isdigit():
        try:
            return str(int(float(s)))
        except Exception:
            pass

    return s


def preparar_base_ccee(df_base):
    """Normaliza colunas de lookup de qualquer base da CCEE."""

    df_base = df_base.copy()

    df_base.columns = [limpar_coluna(c) for c in df_base.columns]

    df_base['CODIGO_CONTRATO'] = (
        df_base['CODIGO_CONTRATO']
        .apply(normalizar_codigo)
    )

    df_base['SIGLA_PERFIL_VENDEDOR'] = (
        df_base['SIGLA_PERFIL_VENDEDOR']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df_base['SIGLA_PERFIL_COMPRADOR'] = (
        df_base['SIGLA_PERFIL_COMPRADOR']
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df_base


def ler_csv_ccee(zip_ref, nome_arquivo):
    """Lê um CSV da CCEE dentro do ZIP (sep=tab, skiprows=1)."""

    with zip_ref.open(nome_arquivo) as f:

        try:
            return pd.read_csv(f, sep='\t', encoding='utf-8', skiprows=1)

        except Exception:
            f.seek(0)
            return pd.read_csv(f, sep='\t', encoding='latin1', skiprows=1)


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
        label='ZIP CCEE',
        type=['zip'],
        key='zip_ccee'
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
# LEITURA ZIP CCEE (cceal_firme + ccear_q)
# =============================================================================

df_cceal_firme = None
df_ccear_q     = None

if arquivo_zip is not None:

    try:

        with zipfile.ZipFile(arquivo_zip) as zip_ref:

            nomes = zip_ref.namelist()

            # cceal_firme — usado para Bismut
            nome_firme = next(
                (n for n in nomes if 'cceal_firme_' in n.lower()),
                None
            )

            if nome_firme:
                df_cceal_firme = preparar_base_ccee(
                    ler_csv_ccee(zip_ref, nome_firme)
                )
                st.success(f'✅ Carregado: {nome_firme}')
            else:
                st.warning('⚠️ cceal_firme não encontrado no ZIP.')

            # ccear_q — usado para Matrix
            nome_ccear_q = next(
                (n for n in nomes if 'ccear_q_' in n.lower()),
                None
            )

            if nome_ccear_q:
                df_ccear_q = preparar_base_ccee(
                    ler_csv_ccee(zip_ref, nome_ccear_q)
                )
                st.success(f'✅ Carregado: {nome_ccear_q}')
            else:
                st.warning('⚠️ ccear_q não encontrado no ZIP.')

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

        st.success('✅ Cliq do mês anterior encontrado!')

    except Exception as erro:

        st.warning(f'⚠️ Erro ao buscar mês anterior: {erro}')


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
# NORMALIZAR COLUNAS DO DF PRINCIPAL PARA O MATCH
# =============================================================================

for col in ['PARTE', 'VENDEDOR', 'COMPRADOR']:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.upper()

for col in ['CLIQ PARADIGMA', 'Cliq Mês Anterior']:
    if col in df.columns:
        df[col] = df[col].apply(normalizar_codigo)


# =============================================================================
# MATCH CLIQ CCEE
#
# Regras:
#   - PARTE contém "BISMUT"  → busca no cceal_firme
#   - CLIQ está em CLIQS_MATRIX → busca no ccear_q
#   - Prioridade: CLIQ PARADIGMA primeiro; Cliq Mês Anterior como fallback
#   - Vendedor e comprador devem bater na base
#   - Se não encontrar: VERIFICAR
#   - Nenhuma regra se aplica: deixa '-'
# =============================================================================

df['CLIQ CCEE'] = '-'

tem_alguma_base = (df_cceal_firme is not None) or (df_ccear_q is not None)

if tem_alguma_base:

    try:

        vazios = {'', '-', 'nan', 'none'}

        def buscar_em_base(df_base, codigo, vendedor, comprador):
            """Retorna True se código + vendedor + comprador existe na base."""

            if df_base is None:
                return False

            resultado = df_base[
                (df_base['CODIGO_CONTRATO']        == codigo)
                &
                (df_base['SIGLA_PERFIL_VENDEDOR']  == vendedor)
                &
                (df_base['SIGLA_PERFIL_COMPRADOR'] == comprador)
            ]

            return not resultado.empty

        def localizar_cliq(linha):

            parte         = str(linha['PARTE']).upper()
            cliq_atual    = linha['CLIQ PARADIGMA']
            cliq_anterior = linha['Cliq Mês Anterior']
            vendedor      = linha['VENDEDOR']
            comprador     = linha['COMPRADOR']

            e_bismut             = 'BISMUT' in parte
            cliq_atual_matrix    = cliq_atual    in CLIQS_MATRIX
            cliq_anterior_matrix = cliq_anterior in CLIQS_MATRIX

            # Se não é Bismut e nenhum CLIQ é Matrix: não processar
            if not e_bismut and not cliq_atual_matrix and not cliq_anterior_matrix:
                return '-'

            # Escolher base para CLIQ PARADIGMA
            if cliq_atual_matrix:
                base_atual = df_ccear_q
            elif e_bismut:
                base_atual = df_cceal_firme
            else:
                base_atual = None

            # Escolher base para Cliq Mês Anterior
            if cliq_anterior_matrix:
                base_anterior = df_ccear_q
            elif e_bismut:
                base_anterior = df_cceal_firme
            else:
                base_anterior = None

            # Tentar CLIQ PARADIGMA (prioridade)
            if cliq_atual.lower() not in vazios and base_atual is not None:

                if buscar_em_base(base_atual, cliq_atual, vendedor, comprador):
                    return cliq_atual

            # Tentar Cliq Mês Anterior (fallback)
            if cliq_anterior.lower() not in vazios and base_anterior is not None:

                if buscar_em_base(base_anterior, cliq_anterior, vendedor, comprador):
                    return cliq_anterior

            return 'VERIFICAR'

        df['CLIQ CCEE'] = df.apply(localizar_cliq, axis=1)

        processadas = df['CLIQ CCEE'] != '-'
        encontrados = processadas & ~df['CLIQ CCEE'].isin(['VERIFICAR'])
        verificar   = df['CLIQ CCEE'] == 'VERIFICAR'

        st.success(
            f'✅ Match CCEE realizado! '
            f'{encontrados.sum()} encontrados | '
            f'{verificar.sum()} para VERIFICAR'
        )

    except Exception as erro:

        st.warning(f'⚠️ Erro no match CCEE: {erro}')


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
