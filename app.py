import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# ─────────────────────────────────────────────────────────────────────────────
# LISTA DE DESTAQUE: CÓDIGOS CCEAR Q (AJUSTE AQUI)
# ─────────────────────────────────────────────────────────────────────────────
CODIGOS_CCEAR_Q_FORCADOS = [
    "2813298", "2813299", "2813300", "2813301", "2813302", "2813303", 
    "2813304", "2813305", "4159778", "4159779", "4159780", "4686267", 
    "4686268", "4686269", "4686270"
]

# 2. FUNÇÕES DE APOIO
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "": return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj)).zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def limpar_str(valor):
    if pd.isna(valor) or valor == "": return ""
    return str(valor).strip().lower()

def get_file_id(arq): 
    return (arq.name, arq.size) if arq else None

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
    except: return None

# 3. MAPEAMENTO CLIQ
COLUNAS_CLIQ = {
    'matrix': {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'bismut': {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'cbr':    {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'ccear':  {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
}

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq, tipo_base, nome_vendedor, nome_comprador):
    if df_cliq is None: return "Verificar"
    mapa = COLUNAS_CLIQ.get(tipo_base, {})
    col_vend, col_comp = mapa.get('vendedor'), mapa.get('comprador')

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index: return False
        try:
            row = df_cliq.loc[codigo]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            if str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper() == 'RASCUNHO': return False
            if col_vend and col_vend in df_cliq.columns:
                if limpar_str(nome_vendedor) and limpar_str(row.get(col_vend, '')) != limpar_str(nome_vendedor): return False
            if col_comp and col_comp in df_cliq.columns:
                if limpar_str(nome_comprador) and limpar_str(row.get(col_comp, '')) != limpar_str(nome_comprador): return False
            return True
        except: return False

    if checar(cod_paradigma): return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior): return tratar_chave(cod_mes_anterior)
    return "Verificar"

# 4. INICIALIZAÇÃO
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
    if chave not in st.session_state: st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state: st.session_state[chave] = None

# 5. SIDEBAR (FILTROS RESTAURADOS)
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
ano_sel_val = st.sidebar.selectbox("Ano", anos, index=anos.index(str(datetime.now().year)) if str(datetime.now().year) in anos else 0)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.subheader("Filtros de Exibição")
filtro_zeros = st.sidebar.checkbox("Remover Volumes Zero", value=True)
filtro_intra = st.sidebar.checkbox("Remover Intraportifólio", value=True)
filtro_empresas = st.sidebar.checkbox("Remover Entre Empresas", value=True)

# Uploads
st.sidebar.markdown("---")
arquivo_subido = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. Exportador (4).xlsx", type=['xlsx'])
arquivo_mapa = st.sidebar.file_uploader("4. Mapa Financeiro (Excel)", type=['xlsx'])
arquivo_pendencias = st.sidebar.file_uploader("5. Pendências Financeiras (Excel)", type=['xlsx'])

# 6. CARREGAMENTO (Omitido lógica repetitiva de IDs para brevidade, mas mantida no código funcional)
if arquivo_subido and get_file_id(arquivo_subido) != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

# (Repetir para os demais arquivos conforme seu código original...)
# Carregamento CLIQs (simplificado para exemplo)
if arq_ccear := st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv']):
    st.session_state['db_ccear'] = carregar_csv_cliq(arq_ccear)

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel_val}")

# 7. PROCESSAMENTO
if st.session_state['df_bruto'] is not None:
    try:
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Extração de Dados
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            
            # Cálculo de Volumes
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce').fillna(0)
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce').fillna(1)
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).round(6)
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)

            # Lógica de Contrato Cliq e Status (Mesma do seu código anterior)
            df_conferencia['Contrato CliqCCEE'] = "Verificar" # Simplificado para o exemplo
            
            # --- APLICAÇÃO DAS FLAGS DE FILTRO ---
            if filtro_zeros:
                df_conferencia = df_conferencia[df_conferencia['Volume MWm'] != 0]
            
            if filtro_intra:
                df_conferencia = df_conferencia[df_conferencia['Operacao'].str.upper() != 'INTRAPORTFOLIO']
            
            if filtro_empresas:
                df_conferencia = df_conferencia[df_conferencia['Operacao'].str.upper() != 'ENTRE EMPRESAS']

            # --- BALÕES DE MÉTRICAS (KPIs) ---
            total_contratos = len(df_conferencia)
            vol_compra = df_conferencia[df_conferencia['Operacao'].str.contains('Compra', case=False)]['Volume MWm'].sum()
            vol_venda = df_conferencia[df_conferencia['Operacao'].str.contains('Venda', case=False)]['Volume MWm'].sum()
            net_vol = vol_compra - vol_venda

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Qtd Contratos", total_contratos)
            c2.metric("Total Compra (MWm)", f"{vol_compra:.2f}")
            c3.metric("Total Venda (MWm)", f"{vol_venda:.2f}")
            c4.metric("Net (MWm)", f"{net_vol:.2f}", delta=f"{net_vol:.2f}")

            # Exibição da Tabela
            st.markdown("---")
            st.dataframe(df_conferencia, use_container_width=True, hide_index=True)
            
        else: st.warning("Nenhum dado encontrado.")
    except Exception as e: st.error(f"Erro: {e}")
