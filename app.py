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

# 3. INICIALIZAÇÃO DO SESSION STATE
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
for chave in ['df_bruto', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias']:
    if chave not in st.session_state: st.session_state[chave] = None if 'df' in chave or 'db' in chave else {}

# 4. INTERFACE LATERAL (RESTAURADA)
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
ano_sel_val = st.sidebar.selectbox("Ano", [str(a) for a in range(2024, 2031)], index=2)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.subheader("Regras de Volume")
aplicar_zerar_intra = st.sidebar.checkbox("Zerar Volume Intraportifólio", value=True)
aplicar_zerar_empresas = st.sidebar.checkbox("Zerar Volume Entre Empresas", value=True)

st.sidebar.markdown("---")
st.sidebar.subheader("Upload das Bases")
arquivo_subido = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. Exportador de Pessoas.xlsx", type=['xlsx'])
arquivo_mapa = st.sidebar.file_uploader("4. Mapa Financeiro.xlsx", type=['xlsx'])
arquivo_pendencias = st.sidebar.file_uploader("5. Pendências Financeiras.xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv'])
arq_cbr = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_matrix = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx', 'csv'])
arq_bismut = st.sidebar.file_uploader("Cliq Bismut", type=['xlsx', 'csv'])

# 5. CARREGAMENTO LÓGICO (SESSION STATE)
if arquivo_subido:
    st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

if arq_ccear: st.session_state['db_ccear'] = carregar_csv_cliq(arq_ccear)
if arq_cbr: st.session_state['db_cbr'] = carregar_csv_cliq(arq_cbr)
if arq_matrix: st.session_state['db_matrix'] = carregar_csv_cliq(arq_matrix)
if arq_bismut: st.session_state['db_bismut'] = carregar_csv_cliq(arq_bismut)

# 6. PROCESSAMENTO
if st.session_state['df_bruto'] is not None:
    try:
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        
        # Filtro por Mês
        df_conferencia = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_conferencia.empty:
            col_boleta = df_base.columns[0]
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            # Mapeamentos
            df_conferencia['Operacao'] = df_conferencia.iloc[:, 1].astype(str)
            df_conferencia['Parte'] = df_conferencia.iloc[:, 62].astype(str).str.strip()
            
            # Cálculo de Volume
            v_mwh = pd.to_numeric(df_conferencia.iloc[:, 20], errors='coerce').fillna(0)
            h_mes = pd.to_numeric(df_conferencia.iloc[:, 15], errors='coerce').fillna(1)
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).round(6)

            # REGRA DE ZERAR VOLUME
            if aplicar_zerar_intra:
                df_conferencia.loc[df_conferencia['Operacao'].str.contains('INTRAPORTFOLIO', case=False, na=False), 'Volume MWm'] = 0
            if aplicar_zerar_empresas:
                df_conferencia.loc[df_conferencia['Operacao'].str.contains('ENTRE EMPRESAS', case=False, na=False), 'Volume MWm'] = 0

            # --- BALÕES DE QUANTIDADE (Métricas) ---
            qtd_compra = len(df_conferencia[df_conferencia['Operacao'].str.contains('Compra', case=False, na=False)])
            qtd_venda = len(df_conferencia[df_conferencia['Operacao'].str.contains('Venda', case=False, na=False)])
            total_ops = len(df_conferencia)

            st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel_val}")
            m1, m2, m3 = st.columns(3)
            m1.metric("Qtd. Operações Compra", f"{qtd_compra}")
            m2.metric("Qtd. Operações Venda", f"{qtd_venda}")
            m3.metric("Total de Operações", f"{total_ops}")
            st.markdown("---")

            # Coluna de Contrato Cliq e Status
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.iloc[:, 60].apply(tratar_chave)
            
            def definir_status(row):
                if row['Contrato CliqCCEE'] in CODIGOS_CCEAR_Q_FORCADOS:
                    return "AJUSTE VALIDADO"
                return "-"
            
            df_conferencia['Status Montante'] = df_conferencia.apply(definir_status, axis=1)

            # Exibição Final
            colunas_finais = [col_boleta, 'Operacao', 'Parte', 'Volume MWm', 'Contrato CliqCCEE', 'Status Montante']
            st.dataframe(df_conferencia[colunas_finais], use_container_width=True, hide_index=True)
            
        else:
            st.warning("Sem dados para este mês.")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
