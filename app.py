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
    except Exception: return None

# 3. INICIALIZAÇÃO DO SESSION STATE (Crucial para evitar o NameError)
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

# Dicionários e Dados
for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias',
              'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
    if chave not in st.session_state: st.session_state[chave] = {} if 'dict' in chave else None

# IDs de Controle de Arquivos
for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state: st.session_state[chave] = None

if 'mes_sel' not in st.session_state: st.session_state['mes_sel'] = meses_nomes[datetime.now().month - 1]
if 'ano_sel' not in st.session_state: st.session_state['ano_sel'] = str(datetime.now().year)

# 4. INTERFACE LATERAL
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=meses_nomes.index(st.session_state['mes_sel']), key='mes_sel')
ano_sel_val = st.sidebar.selectbox("Ano", anos, index=anos.index(st.session_state['ano_sel']) if st.session_state['ano_sel'] in anos else 0, key='ano_sel')
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.markdown("---")
arquivo_subido    = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior  = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx",      type=['xlsx'])
arquivo_pessoas   = st.sidebar.file_uploader("3. Exportador (4).xlsx",          type=['xlsx'])
arquivo_mapa      = st.sidebar.file_uploader("4. Mapa Financeiro (Excel)",     type=['xlsx'])
arquivo_pendencias = st.sidebar.file_uploader("5. Pendências Financeiras (Excel)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear, arq_cbr = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv']), st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1, arq_cceal2 = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx', 'csv']), st.sidebar.file_uploader("Cliq Bismut", type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel_val}")

# 5. CARREGAMENTO E PROCESSAMENTO DAS BASES
# (Lógica de IDs para não reprocessar o Excel a cada clique)

if get_file_id(arquivo_subido) != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    if arquivo_subido:
        st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

if get_file_id(arquivo_pendencias) != st.session_state['fid_pendencias']:
    st.session_state['fid_pendencias'] = get_file_id(arquivo_pendencias)
    if arquivo_pendencias:
        try:
            df_p = pd.read_excel(arquivo_pendencias)
            # Col E (índice 4) Razão Social | Col I (índice 8) Valor
            df_p_clean = df_p.iloc[:, [4, 8]].copy()
            df_p_clean.columns = ['razao_raw', 'valor']
            df_p_clean['valor'] = pd.to_numeric(df_p_clean['valor'], errors='coerce').fillna(0)
            df_p_clean['key'] = df_p_clean['razao_raw'].apply(limpar_str)
            
            # SOMA TODAS AS OCORRÊNCIAS DA MESMA EMPRESA
            df_sum = df_p_clean.groupby('key')['valor'].sum().reset_index()
            st.session_state['dict_pendencias'] = dict(zip(df_sum['key'], df_sum['valor']))
            st.toast("Pendências financeiras carregadas!")
        except: st.session_state['dict_pendencias'] = {}

# [Repetir lógica similar para os outros arquivos conforme necessário...]
# Carregamento simplificado para exemplo (mantendo sua lógica anterior)
if get_file_id(arquivo_mapa) != st.session_state['fid_mapa']:
    st.session_state['fid_mapa'] = get_file_id(arquivo_mapa)
    if arquivo_mapa:
        df_m = pd.read_excel(arquivo_mapa)
        st.session_state['dict_mapa'] = pd.Series(df_m['Situacao_ERP'].values, index=df_m['Codigo_WBC'].apply(tratar_chave).values).to_dict()

# 6. MONTAGEM DA TABELA FINAL
if st.session_state['df_bruto'] is not None:
    df_filtrada = st.session_state['df_bruto'].copy()
    col_mes = df_filtrada.columns[14]
    df_filtrada[col_mes] = pd.to_numeric(df_filtrada[col_mes], errors='coerce')
    df_filtrada = df_filtrada[df_filtrada[col_mes] == mes_num_sel]

    if not df_filtrada.empty:
        col_boleta = df_filtrada.columns[0]
        # Aqui você monta o df_conferencia com as colunas que já tínhamos
        # ... (Mantendo a lógica de colunas que você já validou)
        df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
        df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)
        
        df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_filtrada.columns[1]])
        df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_filtrada.columns[62]]).astype(str).str.strip()
        df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_filtrada.columns[2]]).astype(str).str.strip()
        df_conferencia['Volume MWm'] = (pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_filtrada.columns[20]]), errors='coerce') / 
                                        pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_filtrada.columns[15]]), errors='coerce')).fillna(0)

        # APLICAÇÃO DA PENDÊNCIA (SOMA)
        df_conferencia['key_busca'] = df_conferencia['Razao Social'].apply(limpar_str)
        df_conferencia['Pendência Financeira'] = df_conferencia['key_busca'].map(st.session_state['dict_pendencias']).fillna(0.0)

        # FILTROS INTERATIVOS
        st.write("### Filtros")
        f1, f2 = st.columns(2)
        op_f = f1.selectbox("Filtrar Operação", ["Todos"] + list(df_conferencia['Operacao'].unique()))
        parte_f = f2.selectbox("Filtrar Parte", ["Todos"] + list(df_conferencia['Parte'].unique()))

        df_final = df_conferencia.copy()
        if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
        if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]

        # 7. BALÕES DE TOTAIS INTERATIVOS
        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        # CONTAGEM de operações
        compras_count = len(df_final[df_final['Operacao'].str.upper().str.contains("COMPRA", na=False)])
        vendas_count = len(df_final[df_final['Operacao'].str.upper().str.contains("VENDA", na=False)])
        
        c1.metric("Qtd. Compras", compras_count)
        c2.metric("Qtd. Vendas", vendas_count)
        c3.metric("Total Boletas na Tela", len(df_final))
        st.markdown("---")

        # EXIBIÇÃO
        st.dataframe(df_final.drop(columns=['key_busca']), use_container_width=True, hide_index=True,
                     column_config={"Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f")})
    else:
        st.info("Nenhum dado encontrado para o mês selecionado.")
