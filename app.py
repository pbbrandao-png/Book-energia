import streamlit as st
import pandas as pd
import unicodedata
import calendar


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

def tratar_fonte(valor):
    mapa = {
        'Incentivada 50%':            'Incentivada-I5',
        'Cogeração Qualificada 50%':  'Incentivada-CQ5',
        'Incentivada 100%':           'Incentivada-I1',
        'Incentivada 0%':             'Incentivada-I0',
    }
    return mapa.get(valor, valor)


def tratar_submercado(valor):
    valor = str(valor).strip().upper()
    mapa = {
        'N':     'NORTE',
        'S':     'SUL',
        'NE':    'NORDESTE',
        'SE/CO': 'SUDESTE',
    }
    return mapa.get(valor, valor)


def calcular_cp_lp(dias):
    if pd.isna(dias):
        return '-'
    return 'LP' if dias > 31 else 'CP'


def tratar_modulacao(valor):
    mapa = {
        'C - Carga':  'CARGA',
        'F - Flat':   'FLAT',
        'DECLARADO':  'DECLARADA',
    }
    return mapa.get(valor, valor)


def horas_mes(mes, ano):
    """
    Retorna as horas do mês considerando anos bissextos.
    Requer mês (1-12) e ano (ex: 2024).
    """
    try:
        mes = int(mes)
        ano = int(ano)
        dias = calendar.monthrange(ano, mes)[1]
        return dias * 24
    except Exception:
        return None


def formatar_numero(valor, casas):
    """Formata número com separadores BR (ponto milhar, vírgula decimal)."""
    try:
        return (
            f"{valor:,.{casas}f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )
    except Exception:
        return valor


# =========================================================
# FUNÇÃO PRINCIPAL DE PROCESSAMENTO
# =========================================================

def processar_contratos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Recebe o DataFrame bruto dos contratos aprovados,
    aplica todos os tratamentos e retorna o DataFrame pronto.
    """

    df = df.copy()

    # ── Padroniza nomes de colunas ────────────────────────
    df.columns = [limpar_coluna(col) for col in df.columns]

    # ── Renomeação ────────────────────────────────────────
    renomear = {
        'PARTE_NOME_FANTASIA':      'PARTE',
        'MOVIMENTACAO':             'OPERACAO',
        'FONTE_CONTRATO':           'FONTE',
        'CODIGO_WBC':               'BOLETA',
        'CONTRAPARTE_NOME_FANTASIA':'CONTRAPARTE',
        'QUANTATUALIZADA':          'MONTANTE_MWH',
        'CODIGO_CCEE':              'CLIQ PARADIGMA',
        'TIPO_DE_MODULACAO':        'MODULACAO WBC',
        'FLEXLIMITE_MODULACAOMAX':  'MOD MAX',
        'FLEXLIMITE_MODULACAOMIN':  'MOD MIN',
    }
    # Só renomeia colunas que existem (evita erro silencioso)
    renomear_valido = {k: v for k, v in renomear.items() if k in df.columns}
    colunas_nao_encontradas = set(renomear.keys()) - set(renomear_valido.keys())
    if colunas_nao_encontradas:
        st.warning(
            f"⚠️ Colunas não encontradas no arquivo e ignoradas: "
            f"{', '.join(sorted(colunas_nao_encontradas))}"
        )
    df = df.rename(columns=renomear_valido)

    # ── Fonte ─────────────────────────────────────────────
    if 'FONTE' in df.columns:
        df['FONTE'] = df['FONTE'].apply(tratar_fonte)

    # ── Submercado ────────────────────────────────────────
    if 'SUBMERCADO' in df.columns:
        df['SUBMERCADO'] = df['SUBMERCADO'].apply(tratar_submercado)

    # ── Modulação ─────────────────────────────────────────
    if 'MODULACAO WBC' in df.columns:
        df['MODULACAO WBC'] = df['MODULACAO WBC'].apply(tratar_modulacao)

    # ── Datas e CP/LP ─────────────────────────────────────
    if 'SUPRIMENTO_INICIO' in df.columns and 'SUPRIMENTO_TERMINO' in df.columns:
        df['SUPRIMENTO_INICIO'] = pd.to_datetime(
            df['SUPRIMENTO_INICIO'], errors='coerce'
        )
        df['SUPRIMENTO_TERMINO'] = pd.to_datetime(
            df['SUPRIMENTO_TERMINO'], errors='coerce'
        )
        df['DIAS'] = (
            df['SUPRIMENTO_TERMINO'] - df['SUPRIMENTO_INICIO']
        ).dt.days
        df['CP/LP'] = df['DIAS'].apply(calcular_cp_lp)

    # ── Horas do mês (com suporte a bissexto) ─────────────
    # Tenta derivar o ano de SUPRIMENTO_INICIO; se não existir, usa o ano atual
    tem_inicio = 'SUPRIMENTO_INICIO' in df.columns
    if 'MES' in df.columns:
        df['MES'] = pd.to_numeric(df['MES'], errors='coerce')

        if tem_inicio:
            df['ANO'] = df['SUPRIMENTO_INICIO'].dt.year
        else:
            df['ANO'] = pd.Timestamp.today().year

        df['HORAS_MES'] = df.apply(
            lambda row: horas_mes(row['MES'], row['ANO']), axis=1
        )

    # ── Montante ─────────────────────────────────────────
    if 'MONTANTE_MWH' in df.columns:
        df['MONTANTE_MWH_NUM'] = pd.to_numeric(df['MONTANTE_MWH'], errors='coerce')

        if 'HORAS_MES' in df.columns:
            df['MONTANTE_MWM_NUM'] = df['MONTANTE_MWH_NUM'] / df['HORAS_MES']
        
        # Colunas formatadas APENAS para exibição (não sobrescrevem as numéricas)
        df['MONTANTE MWh'] = df['MONTANTE_MWH_NUM'].apply(
            lambda v: formatar_numero(v, 3)
        )
        if 'MONTANTE_MWM_NUM' in df.columns:
            df['MONTANTE MWm'] = df['MONTANTE_MWM_NUM'].apply(
                lambda v: formatar_numero(v, 6)
            )

    return df


# =========================================================
# INTERFACE
# =========================================================

st.set_page_config(page_title="Book de Energia", layout="wide")
st.title("⚡ Book de Energia")

col1, col2 = st.columns(2)
with col1:
    arquivo = st.file_uploader(
        "Contratos aprovados",
        type=['xlsx', 'csv', 'xlsm'],
        key='aprovados'
    )
with col2:
    arquivo_2 = st.file_uploader(
        "Contratos mês anterior",
        type=['xlsx', 'csv', 'xlsm'],
        key='mes_anterior'
    )


# =========================================================
# PROCESSAMENTO
# =========================================================

if arquivo is None:
    st.info("📂 Faça o upload do arquivo de contratos aprovados para começar.")
    st.stop()

st.success("✅ Arquivo carregado com sucesso!")

# ── Leitura ───────────────────────────────────────────────
try:
    df_contratos_aprovados = pd.read_excel(arquivo, skiprows=8)
except Exception as e:
    st.error(f"❌ Erro ao ler o arquivo de contratos aprovados: {e}")
    st.stop()

df_contratos_aprovados = df_contratos_aprovados.fillna("-")

df_mes_anterior = None
if arquivo_2 is not None:
    try:
        df_mes_anterior = pd.read_excel(arquivo_2)
        df_mes_anterior = df_mes_anterior.fillna("-")
        st.success("✅ Arquivo do mês anterior carregado com sucesso!")
    except Exception as e:
        st.warning(f"⚠️ Erro ao ler o arquivo do mês anterior: {e}")

# ── Processamento ─────────────────────────────────────────
try:
    df_tratado = processar_contratos(df_contratos_aprovados)
except Exception as e:
    st.error(f"❌ Erro durante o processamento: {e}")
    st.stop()

# ── Debug (colunas disponíveis) ───────────────────────────
with st.expander("🔍 Colunas encontradas no arquivo"):
    st.write(df_tratado.columns.tolist())

# ── Seleção das colunas para exibição ────────────────────
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
    'MOD MAX',
]

colunas_existentes = [col for col in colunas_desejadas if col in df_tratado.columns]

colunas_faltando = set(colunas_desejadas) - set(colunas_existentes)
if colunas_faltando:
    st.warning(
        f"⚠️ Colunas esperadas não encontradas na exibição: "
        f"{', '.join(sorted(colunas_faltando))}"
    )

# ── Exibição ──────────────────────────────────────────────
st.subheader("Contratos Aprovados")
st.dataframe(
    df_tratado[colunas_existentes],
    hide_index=True,
    use_container_width=True,
)

# ── Placeholder para lógica do mês anterior ──────────────
if df_mes_anterior is not None:
    st.subheader("Contratos Mês Anterior")
    st.dataframe(df_mes_anterior, hide_index=True, use_container_width=True)
    # TODO: implementar comparação entre df_tratado e df_mes_anterior
