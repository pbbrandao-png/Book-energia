import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURACAO DA PAGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNCOES DE APOIO
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
    if arquivo is None:
        return None
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
    except Exception:
        return None

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq):
    if df_cliq is None:
        return "Verificar"

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index:
            return False
        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        situacao = str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper()
        return situacao != 'RASCUNHO'

    if checar(cod_paradigma):
        return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior):
        return tratar_chave(cod_mes_anterior)
    return "Verificar"

# ---------------------------------------------------------------------------
# 3. INICIALIZA SESSION STATE
# ---------------------------------------------------------------------------
meses_nomes = [
    "Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
anos = [str(a) for a in range(2024, 2031)]

if 'mes_sel' not in st.session_state:
    st.session_state['mes_sel'] = meses_nomes[datetime.now().month - 1]
if 'ano_sel' not in st.session_state:
    st.session_state['ano_sel'] = str(datetime.now().year)

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor',
              'db_matrix', 'db_bismut']:
    if chave not in st.session_state:
        st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas',
              'chave_matrix', 'fid_cceal2']:
    if chave not in st.session_state:
        st.session_state[chave] = None

# ---------------------------------------------------------------------------
# 4. INTERFACE LATERAL
# ---------------------------------------------------------------------------
st.sidebar.title("Configurações")
st.sidebar.subheader("Período de Referência")

mes_nome_sel = st.sidebar.selectbox(
    "Mês", 
    meses_nomes, 
    index=meses_nomes.index(st.session_state['mes_sel']),
    key='mes_sel'
)

ano_sel = st.sidebar.selectbox(
    "Ano", 
    anos, 
    index=anos.index(st.session_state['ano_sel']) if st.session_state['ano_sel'] in anos else 0,
    key='ano_sel'
)

mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.markdown("---")

def get_file_id(arq):
    if arq is None: return None
    return (arq.name, arq.size)

arquivo_subido   = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx",      type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. Exportador (4).xlsx",          type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q",    type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq Matrix",      type=['xlsx', 'csv'])
arq_cceal2 = st.sidebar.file_uploader("Cliq Bismut",      type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

# ---------------------------------------------------------------------------
# 5. CARREGAMENTO DOS DADOS
# ---------------------------------------------------------------------------
fid_subido = get_file_id(arquivo_subido)
if fid_subido != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = fid_subido
    if arquivo_subido:
        try:
            st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        except Exception as e:
            st.session_state['df_bruto'] = None
            st.error(f"Erro ao carregar base principal: {e}")

fid_anterior = get_file_id(arquivo_anterior)
if fid_anterior != st.session_state['fid_anterior']:
    st.session_state['fid_anterior'] = fid_anterior
    dict_ma = {}
    if arquivo_anterior:
        try:
            df_apoio = pd.read_excel(arquivo_anterior, header=0, dtype=str)
            df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].apply(tratar_chave)
            df_apoio.iloc[:, 1] = df_apoio.iloc[:, 1].apply(tratar_chave)
            dict_ma = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()
        except Exception: pass
    st.session_state['dict_mes_anterior'] = dict_ma

fid_pessoas = get_file_id(arquivo_pessoas)
if fid_pessoas != st.session_state['fid_pessoas']:
    st.session_state['fid_pessoas'] = fid_pessoas
    dict_v, dict_c = {}, {}
    if arquivo_pessoas:
        try:
            df_pers = pd.read_excel(arquivo_pessoas)
            df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
            dict_c = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
            dict_v = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
        except Exception: pass
    st.session_state['dict_comprador'] = dict_c
    st.session_state['dict_vendedor']  = dict_v

fid_ccear, fid_cbr, fid_cceal1 = get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)
chave_matrix = (fid_ccear, fid_cbr, fid_cceal1)
if chave_matrix != st.session_state['chave_matrix']:
    st.session_state['chave_matrix'] = chave_matrix
    dfs = []
    for arq in [arq_ccear, arq_cbr, arq_cceal1]:
        df = carregar_csv_cliq(arq)
        if df is not None: dfs.append(df)
    st.session_state['db_matrix'] = pd.concat(dfs) if dfs else None

fid_cceal2 = get_file_id(arq_cceal2)
if fid_cceal2 != st.session_state['fid_cceal2']:
    st.session_state['fid_cceal2'] = fid_cceal2
    st.session_state['db_bismut'] = carregar_csv_cliq(arq_cceal2)

# Atalhos para os dados
df_bruto = st.session_state['df_bruto']
dict_mes_anterior = st.session_state['dict_mes_anterior']
dict_comprador = st.session_state['dict_comprador']
dict_vendedor = st.session_state['dict_vendedor']
db_matrix = st.session_state['db_matrix']
db_bismut = st.session_state['db_bismut']

# ---------------------------------------------------------------------------
# 6. PROCESSAMENTO
# ---------------------------------------------------------------------------
if df_bruto is not None:
    try:
        # Mapeamento de Colunas por Índice
        col_boleta         = df_bruto.columns[0]
        col_operacao       = df_bruto.columns[1]
        col_cnpj           = df_bruto.columns[4]
        col_submercado_raw = df_bruto.columns[8]  # Coluna I (Indice 8)
        col_energia        = df_bruto.columns[5]
        col_contraparte    = df_bruto.columns[6]
        col_cp_lp          = df_bruto.columns[12] # Coluna M (Indice 12)
        col_mes_suprimento = df_bruto.columns[14]
        col_horas_mes      = df_bruto.columns[15]
        col_montante_mwh   = df_bruto.columns[17]
        col_volume_mwh     = df_bruto.columns[20]
        col_mod_min        = df_bruto.columns[28]
        col_mod_max        = df_bruto.columns[29]
        col_cliq_para      = df_bruto.columns[60]
        col_parte_bk       = df_bruto.columns[62]
        col_mod_wbc        = df_bruto.columns[63]

        df_num = df_bruto.copy()
        df_num[col_mes_suprimento] = pd.to_numeric(df_num[col_mes_suprimento], errors='coerce')
        df_filtrada = df_num[df_num[col_mes_suprimento] == mes_num_sel].copy()

        if not df_filtrada.empty:
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Mapeamentos Diretos
            df_conferencia['Operacao']         = df_conferencia[col_boleta].map(df_lookup[col_operacao]).astype(str)
            df_conferencia['Tipo Energia']     = df_conferencia[col_boleta].map(df_lookup[col_energia]).astype(str).str.strip()
            df_conferencia['Parte']            = df_conferencia[col_boleta].map(df_lookup[col_parte_bk]).astype(str).str.strip()
            
            # Lógica Especial Energia
            mask_jaca = df_conferencia['Parte'].str.upper().str.contains('UFV JACARANDA 1', na=False)
            df_conferencia.loc[mask_jaca, 'Tipo Energia'] = 'Incentivada-I5'
            
            mapa_energia = {
                'INCENTIVADA-50%': 'Incentivada-I5', 
                'INCENTIVADA-CQ50%': 'Incentivada-CQ5', 
                'INCENTIVADA-100%': 'Incentivada-I1', 
                'INCENTIVADA-0%': 'Incentivada-I0'
            }
            df_conferencia['Tipo Energia'] = df_conferencia['Tipo Energia'].apply(lambda v: mapa_energia.get(str(v).strip().upper(), v))
            
            # Dados de Contraparte e CP/LP
            df_conferencia['Contraparte']      = df_conferencia[col_boleta].map(df_lookup[col_contraparte])
            df_conferencia['CP/LP']            = df_conferencia[col_boleta].map(df_lookup[col_cp_lp])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cnpj]).apply(formatar_cnpj)
            
            # Lógica de Submercado (Coluna I)
            mapa_submercado = {
                'SE/CO': 'Sudeste',
                'N': 'Norte',
                'NE': 'Nordeste',
                'S': 'Sul'
            }
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[col_submercado_raw]).replace(mapa_submercado)
            
            # Cálculos de Volume
            df_conferencia['Montante MWh']     = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_montante_mwh]), errors='coerce').fillna(0).round(3)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_volume_mwh]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_horas_mes]), errors='coerce')
            df_conferencia['Volume MWm']       = (v_mwh / h_mes).fillna(0).round(6)

            # Modulações e Cliq
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[col_cliq_para]).apply(tratar_chave)
            df_conferencia['Modulacao WBC']      = df_conferencia[col_boleta].map(df_lookup[col_mod_wbc]).apply(limpar_modulacao)
            df_conferencia['Modulacao Minima']   = df_conferencia[col_boleta].map(df_lookup[col_mod_min])
            df_conferencia['Modulacao Maxima']   = df_conferencia[col_boleta].map(df_lookup[col_mod_max])

            # De-Para de outras bases
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("N/A")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("N/A")

            def resolver_cliq(row):
                parte = str(row['Parte']).strip().upper()
                db = db_bismut if 'BISMUT' in parte else db_matrix
                return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], db)

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)

            # Ordenação numérica por boleta
            df_conferencia['_boleta_num'] = pd.to_numeric(df_conferencia['Boleta_Key'], errors='coerce')
            df_conferencia = df_conferencia.sort_values('_boleta_num').drop(columns=['_boleta_num'])

            # Filtros na UI
            st.write("### Filtros Rápidos")
            f1, f2, f3 = st.columns(3)
            op_f = f1.selectbox("Operação", ["Todos"] + sorted(list(df_conferencia['Operacao'].unique())))
            parte_f = f2.selectbox("Parte", ["Todos"] + sorted(list(df_conferencia['Parte'].unique())))
            rem_zero = f3.checkbox("Ocultar Zerados", value=False)

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if rem_zero: df_final = df_final[df_final['Volume MWm'] != 0]

            # Definição da ordem das colunas para exibição
            ordem_final = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP',
                'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 
                'CliqCCEE Paradigma', 'Modulacao WBC', 'Modulacao Minima', 
                'Modulacao Maxima', 'Contrato CliqCCEE mes anterior', 
                'Comprador', 'Vendedor', 'Contrato CliqCCEE'
            ]
            
            st.dataframe(
                df_final[ordem_final], 
                hide_index=True, 
                use_container_width=True,
                column_config={
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"),
                    "Volume MWm": st.column_config.NumberColumn(format="%.6f")
                }
            )
        else:
            st.warning(f"Nenhum dado encontrado para {mes_nome_sel}/{ano_sel}")

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
