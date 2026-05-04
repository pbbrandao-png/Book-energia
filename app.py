import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES COM CACHE (Para evitar lentidão)
@st.cache_data
def carregar_csv_cliq_cached(arquivo_lista):
    """Lê e concatena as bases CCEE apenas quando os arquivos mudam."""
    dfs = []
    for arquivo in arquivo_lista:
        if arquivo is not None:
            try:
                nome = arquivo.name
                if nome.endswith('.csv'):
                    df = pd.read_csv(arquivo, sep='\t', encoding='latin-1', skiprows=1, dtype=str)
                else:
                    df = pd.read_excel(arquivo, dtype=str)
                
                if 'CODIGO_CONTRATO' in df.columns:
                    df['CODIGO_CONTRATO'] = df['CODIGO_CONTRATO'].str.strip().str.replace('.0', '', regex=False)
                    df = df.set_index('CODIGO_CONTRATO')
                    dfs.append(df)
            except:
                continue
    return pd.concat(dfs) if dfs else None

def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "": return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj)).zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

# Seleção de período sem atualização automática (manual via botão)
meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
# Defini o padrão como Abril (index 3) para não iniciar em Maio
mes_nome_sel = st.sidebar.selectbox("Mês de Referência", meses_nomes, index=3)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1
ano_sel = st.sidebar.selectbox("Ano", [str(a) for a in range(2024, 2031)], index=2)

st.sidebar.markdown("---")
arquivo_subido   = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior", type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. RelPers_858 (Pessoas)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq CCEAL Firme 101457", type=['xlsx', 'csv'])
arq_cceal2 = st.sidebar.file_uploader("Cliq CCEAL Firme 101475 (Bismut)", type=['xlsx', 'csv'])

# BOTÃO DE PROCESSAMENTO (Crucial para não travar)
processar = st.sidebar.button("🚀 Processar Dados", use_container_width=True)

# 4. EXECUÇÃO
st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

if processar and arquivo_subido:
    with st.spinner('Processando bases pesadas...'):
        try:
            # Carregamento cacheado das bases CCEE
            db_matrix = carregar_csv_cliq_cached([arq_ccear, arq_cbr, arq_cceal1])
            db_bismut = carregar_csv_cliq_cached([arq_cceal2])

            # Processamento da base principal
            df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
            
            # Pegando as colunas por índice para evitar erro de nome
            col_boleta = df_bruto.columns[0]
            col_mes = df_bruto.columns[14]
            
            df_bruto[col_mes] = pd.to_numeric(df_bruto[col_mes], errors='coerce')
            df_filtrada = df_bruto[df_bruto[col_mes] == mes_num_sel].copy()

            if df_filtrada.empty:
                st.warning(f"Sem dados para o mês {mes_num_sel}.")
            else:
                # Dicionários de apoio
                dict_mes_ant = {}
                if arquivo_anterior:
                    df_ant = pd.read_excel(arquivo_anterior, dtype=str)
                    dict_mes_ant = pd.Series(df_ant.iloc[:, 1].values, index=df_ant.iloc[:, 0].apply(tratar_chave).values).to_dict()

                # Construção do resultado
                df_res = pd.DataFrame()
                df_res['Boleta'] = df_filtrada[col_boleta].apply(tratar_chave)
                df_res['Parte'] = df_filtrada[df_bruto.columns[62]].astype(str).str.strip()
                df_res['Volume MWm'] = (df_filtrada[df_bruto.columns[20]] / df_filtrada[df_bruto.columns[15]]).round(4)
                df_res['Paradigma'] = df_filtrada[df_bruto.columns[60]].apply(tratar_chave)
                df_res['Mês Anterior'] = df_res['Boleta'].map(dict_mes_ant).fillna("-")

                def validar_cliq(row):
                    db = db_bismut if 'BISMUT' in str(row['Parte']).upper() else db_matrix
                    if db is None: return "Verificar"
                    
                    for cod in [row['Paradigma'], row['Mês Anterior']]:
                        c = tratar_chave(cod)
                        if c in db.index:
                            # CORREÇÃO DO ERRO: SITUACAO_CONTRATO
                            status = str(db.loc[c, 'SITUACAO_CONTRATO']).upper() if 'SITUACAO_CONTRATO' in db.columns else ""
                            if "RASCUNHO" not in status: return c
                    return "Verificar"

                df_res['Contrato CliqCCEE'] = df_res.apply(validar_cliq, axis=1)

                # --- EXIBIÇÃO ---
                pendentes = len(df_res[df_res['Contrato CliqCCEE'] == "Verificar"])
                st.metric("Contratos Pendentes", pendentes)

                # CORREÇÃO DO ERRO: .map() em vez de .applymap() para o Styler
                def highlight_err(s):
                    return ['color: red; font-weight: bold' if v == "Verificar" else '' for v in s]

                st.dataframe(df_res.style.apply(highlight_err, subset=['Contrato CliqCCEE']), use_container_width=True)

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
else:
    st.info("Aguardando clique no botão 'Processar Dados' para iniciar.")
