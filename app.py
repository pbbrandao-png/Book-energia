import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES DE APOIO
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "": return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj)).zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def limpar_modulacao(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper()
    if "FLAT" in t: return "Flat"
    if "CARGA" in t: return "Carga"
    if "DECLARADO" in t or "INFORMADO" in t: return "Declarado"
    if "GERA" in t: return "Geracao"
    return texto

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def carregar_csv_cliq(arquivo):
    if arquivo is None: return None
    try:
        nome = arquivo.name if hasattr(arquivo, 'name') else str(arquivo)
        if nome.endswith('.csv'):
            df = pd.read_csv(arquivo, sep='\t', encoding='latin-1', skiprows=1, dtype=str)
        else:
            df = pd.read_excel(arquivo, dtype=str)
        
        if 'CODIGO_CONTRATO' in df.columns:
            df['CODIGO_CONTRATO'] = df['CODIGO_CONTRATO'].apply(tratar_chave)
            df = df.set_index('CODIGO_CONTRATO')
            return df
        return None
    except Exception: return None

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq):
    if df_cliq is None: return "Verificar"
    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index: return False
        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame): row = row.iloc[0]
        situacao = str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper()
        return situacao != 'RASCUNHO'
    if checar(cod_paradigma): return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior): return tratar_chave(cod_mes_anterior)
    return "Verificar"

# ─────────────────────────────────────────────────────────────────────────────
# COMO ADICIONAR NOVAS COLUNAS VIA LOOKUP (dicionário de planilha)
# ─────────────────────────────────────────────────────────────────────────────
# Padrão de 3 passos:
#
#  PASSO 1 — Carregar (na seção "fid_pessoas"):
#    Leia a coluna desejada e crie um dicionário  {chave: valor}
#    onde a chave é sempre a coluna D (índice 3) tratada com tratar_chave().
#    Exemplo:
#      dict_nova = pd.Series(df_pers.iloc[:, INDICE_COLUNA].values,
#                            index=df_pers['chave'].values).to_dict()
#    Salve em st.session_state['dict_nova'].
#
#  PASSO 2 — Buscar no processamento:
#    df_conferencia['Nome da Coluna'] = (
#        df_conferencia['Boleta_Key']          # chave de ligação
#        .map(st.session_state['dict_nova'])   # faz o lookup
#        .fillna("-")                           # valor padrão se não achar
#    )
#
#  PASSO 3 — Incluir na lista `ordem`:
#    Adicione 'Nome da Coluna' na posição desejada dentro da lista `ordem`.
#
# Resumo de índices do Exportador (4).xlsx:
#   Coluna B (índice 1) → Comprador    → dict_comprador
#   Coluna C (índice 2) → Vendedor     → dict_vendedor
#   Coluna D (índice 3) → Chave (ref)  → usada como índice
# ─────────────────────────────────────────────────────────────────────────────

# 3. INICIALIZA SESSION STATE
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

if 'mes_sel' not in st.session_state: st.session_state['mes_sel'] = meses_nomes[datetime.now().month - 1]
if 'ano_sel' not in st.session_state: st.session_state['ano_sel'] = str(datetime.now().year)

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'db_matrix', 'db_bismut']:
    if chave not in st.session_state: st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2']:
    if chave not in st.session_state: st.session_state[chave] = None

# 4. INTERFACE LATERAL
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=meses_nomes.index(st.session_state['mes_sel']), key='mes_sel')
ano_sel = st.sidebar.selectbox("Ano", anos, index=anos.index(st.session_state['ano_sel']) if st.session_state['ano_sel'] in anos else 0, key='ano_sel')
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.markdown("---")
def get_file_id(arq): return (arq.name, arq.size) if arq else None

arquivo_subido = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. Exportador (4).xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q",    type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq Matrix",     type=['xlsx', 'csv'])
arq_cceal2 = st.sidebar.file_uploader("Cliq Bismut",     type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

# 5. CARREGAMENTO DOS DADOS

# ── Contratos aprovados ──────────────────────────────────────────────────────
fid_subido = get_file_id(arquivo_subido)
if fid_subido != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = fid_subido
    if arquivo_subido:
        try: st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        except: st.session_state['df_bruto'] = None

# ── Base mês anterior ────────────────────────────────────────────────────────
fid_anterior = get_file_id(arquivo_anterior)
if fid_anterior != st.session_state['fid_anterior']:
    st.session_state['fid_anterior'] = fid_anterior
    dict_ma = {}
    if arquivo_anterior:
        try:
            df_apoio = pd.read_excel(arquivo_anterior, dtype=str)
            dict_ma = pd.Series(
                df_apoio.iloc[:, 1].values,
                index=df_apoio.iloc[:, 0].apply(tratar_chave).values
            ).to_dict()
        except: pass
    st.session_state['dict_mes_anterior'] = dict_ma

# ── Exportador (4).xlsx — Comprador (col B) e Vendedor (col C) ───────────────
# Coluna D (índice 3) é a chave de ligação com a boleta.
# Se quiser adicionar outra coluna do mesmo arquivo, siga o padrão:
#   dict_nova = pd.Series(df_pers.iloc[:, ÍNDICE].values, index=df_pers['chave'].values).to_dict()
#   st.session_state['dict_nova'] = dict_nova
fid_pessoas = get_file_id(arquivo_pessoas)
if fid_pessoas != st.session_state['fid_pessoas']:
    st.session_state['fid_pessoas'] = fid_pessoas
    dict_v, dict_c = {}, {}
    if arquivo_pessoas:
        try:
            df_pers = pd.read_excel(arquivo_pessoas)
            # Coluna D (índice 3) → chave de busca
            df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
            # Coluna B (índice 1) → Comprador
            dict_c = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
            # Coluna C (índice 2) → Vendedor
            dict_v = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
        except: pass
    st.session_state['dict_comprador'] = dict_c
    st.session_state['dict_vendedor']  = dict_v

# ── Cliq CCEE (Matrix + Bismut) ──────────────────────────────────────────────
if (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)) != st.session_state['chave_matrix']:
    st.session_state['chave_matrix'] = (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1))
    dfs = [carregar_csv_cliq(a) for a in [arq_ccear, arq_cbr, arq_cceal1] if a]
    st.session_state['db_matrix'] = pd.concat(dfs) if dfs else None

if get_file_id(arq_cceal2) != st.session_state['fid_cceal2']:
    st.session_state['fid_cceal2'] = get_file_id(arq_cceal2)
    st.session_state['db_bismut'] = carregar_csv_cliq(arq_cceal2)

# 6. PROCESSAMENTO
df_bruto = st.session_state['df_bruto']
if df_bruto is not None:
    try:
        df_num = df_bruto.copy()
        col_mes = df_bruto.columns[14]
        df_num[col_mes] = pd.to_numeric(df_num[col_mes], errors='coerce')
        df_filtrada = df_num[df_num[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_bruto.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # ── Colunas Base ──────────────────────────────────────────────────
            df_conferencia['Operacao']        = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[1]]).astype(str)
            df_conferencia['Tipo Energia']    = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[5]]).astype(str).str.strip()
            df_conferencia['Parte']           = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[62]]).astype(str).str.strip()
            df_conferencia['Contraparte']     = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[6]])
            df_conferencia['CP/LP']           = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[12]])
            df_conferencia['CNPJ Contraparte']= df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[4]]).apply(formatar_cnpj)

            mapa_sub = {'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'}
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[8]]).replace(mapa_sub)

            df_conferencia['Montante MWh'] = pd.to_numeric(
                df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[17]]), errors='coerce'
            ).fillna(0).round(3)

            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)

            # ── Modulações e Cliq ─────────────────────────────────────────────
            df_conferencia['CliqCCEE Paradigma']          = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC']               = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[63]]).apply(limpar_modulacao)
            df_conferencia['Modulacao Minima']            = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[28]])
            df_conferencia['Modulacao Maxima']            = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[29]])
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")

            def resolver(row):
                db = st.session_state['db_bismut'] if 'BISMUT' in str(row['Parte']).upper() else st.session_state['db_matrix']
                return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], db)
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver, axis=1)

            # ── PASSO 2: buscar Comprador e Vendedor via Boleta_Key ───────────
            # A Boleta_Key é a chave de ligação com a coluna D do Exportador.
            # .map() percorre o dicionário e retorna o valor correspondente.
            # .fillna("-") garante que linhas sem correspondência fiquem com "-".
            df_conferencia['Comprador'] = (
                df_conferencia['Boleta_Key']
                .map(st.session_state['dict_comprador'])
                .fillna("-")
            )
            df_conferencia['Vendedor'] = (
                df_conferencia['Boleta_Key']
                .map(st.session_state['dict_vendedor'])
                .fillna("-")
            )

            # ── Filtros ───────────────────────────────────────────────────────
            st.write("### Filtros")
            f1, f2, f3, f4 = st.columns([2, 2, 2, 1])
            op_f    = f1.selectbox("Operação",          ["Todos"] + sorted(df_conferencia['Operacao'].dropna().unique().tolist()))
            parte_f = f2.selectbox("Parte",             ["Todos"] + sorted(df_conferencia['Parte'].dropna().unique().tolist()))
            cliq_f  = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].dropna().unique().tolist()))
            rem_zero = f4.toggle("Ocultar Zero", value=False)

            df_final = df_conferencia.copy()
            if op_f    != "Todos": df_final = df_final[df_final['Operacao']         == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte']            == parte_f]
            if cliq_f  != "Todos": df_final = df_final[df_final['Contrato CliqCCEE']== cliq_f]
            if rem_zero:           df_final = df_final[df_final['Volume MWm']       != 0]

            # ── Resumo ────────────────────────────────────────────────────────
            st.write("### Resumo")
            m1, m2, m3 = st.columns(3)
            m1.metric("Contratos", len(df_final))
            m2.metric("Compras",   len(df_final[df_final['Operacao'].str.upper() == 'COMPRA']))
            m3.metric("Vendas",    len(df_final[df_final['Operacao'].str.upper() == 'VENDA']))
            st.markdown("---")

            # ── PASSO 3: ordem das colunas (Comprador e Vendedor incluídos) ───
            ordem = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte',
                'Contraparte', 'CP/LP', 'CNPJ Contraparte',
                'Submercado', 'Montante MWh', 'Volume MWm',
                'CliqCCEE Paradigma',
                'Modulacao WBC', 'Modulacao Minima', 'Modulacao Maxima',
                'Contrato CliqCCEE mes anterior','Comprador', 'Vendedor', 'Contrato CliqCCEE',
            ]

            st.dataframe(
                df_final[ordem].sort_values(by=col_boleta),
                hide_index=True, use_container_width=True,
                column_config={
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"),
                    "Volume MWm":   st.column_config.NumberColumn(format="%.6f"),
                }
            )
        else:
            st.warning(f"Sem dados para {mes_nome_sel}/{ano_sel}")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
