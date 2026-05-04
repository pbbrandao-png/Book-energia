import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES DE APOIO E TRATAMENTO
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
    except Exception:
        return None

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq):
    """
    Valida o código do contrato:
    1. Prioridade para o Paradigma.
    2. Secundário para o Mês Anterior.
    3. Valida se SITUACAO_CONTRATO != 'RASCUNHO'.
    """
    if df_cliq is None: return "Verificar"

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index:
            return False
        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        situacao = str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper()
        return situacao != 'RASCUNHO'

    if checar(cod_paradigma): return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior): return tratar_chave(cod_mes_anterior)
    return "Verificar"

# 3. INTERFACE LATERAL (INPUTS)
st.sidebar.title("Configurações")

st.sidebar.subheader("Período de Referência")
meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1
ano_sel = st.sidebar.selectbox("Ano", [str(a) for a in range(2024, 2031)], index=2)

st.sidebar.markdown("---")
arquivo_subido   = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior", type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. RelPers_858 (Pessoas)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE - Matrix/CBR/CCEAR")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq CCEAL Firme 101457", type=['xlsx', 'csv'])

st.sidebar.subheader("Bases Cliq CCEE - Bismut")
arq_cceal2 = st.sidebar.file_uploader("Cliq CCEAL Firme 101475", type=['xlsx', 'csv'])

# 4. CARREGAMENTO E CONSOLIDAÇÃO
# Bases Matrix
dfs_matrix = []
for arq in [arq_ccear, arq_cbr, arq_cceal1]:
    df = carregar_csv_cliq(arq)
    if df is not None: dfs_matrix.append(df)
db_matrix = pd.concat(dfs_matrix) if dfs_matrix else None

# Base Bismut
db_bismut = carregar_csv_cliq(arq_cceal2)

# Dicionários de Apoio
dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_ant = pd.read_excel(arquivo_anterior, dtype=str)
        df_ant.iloc[:, 0] = df_ant.iloc[:, 0].apply(tratar_chave)
        df_ant.iloc[:, 1] = df_ant.iloc[:, 1].apply(tratar_chave)
        dict_mes_anterior = pd.Series(df_ant.iloc[:, 1].values, index=df_ant.iloc[:, 0].values).to_dict()
    except: pass

dict_vendedor, dict_comprador = {}, {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor  = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
    except: pass

# 5. PROCESSAMENTO
st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de Colunas
        col_boleta = df_bruto.columns[0]; col_oper = df_bruto.columns[1]
        col_cnpj = df_bruto.columns[4]; col_cp = df_bruto.columns[6]
        col_parte = df_bruto.columns[62]; col_mes = df_bruto.columns[14]
        col_horas = df_bruto.columns[15]; col_vol = df_bruto.columns[20]
        col_cliq_p = df_bruto.columns[60]; col_mod_w = df_bruto.columns[63]

        df_bruto[col_mes] = pd.to_numeric(df_bruto[col_mes], errors='coerce')
        df_filtrada = df_bruto[df_bruto[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            # Construção do DataFrame de Conferência
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[col_oper]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[col_parte]).astype(str).str.strip()
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cp])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cnpj]).apply(formatar_cnpj)
            
            v_mwh = df_conferencia[col_boleta].map(df_lookup[col_vol])
            h_mes = df_conferencia[col_boleta].map(df_lookup[col_horas])
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(4)

            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[col_cliq_p]).apply(tratar_chave)
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")
            
            # Lógica de Resolução de Contrato
            def resolver(row):
                db = db_bismut if 'BISMUT' in str(row['Parte']).upper() else db_matrix
                return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], db)

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver, axis=1)
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("N/A")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("N/A")
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[col_mod_w]).apply(limpar_modulacao)

            # --- INTERFACE DE RESULTADOS ---
            pendentes = len(df_conferencia[df_conferencia['Contrato CliqCCEE'] == "Verificar"])
            st.metric("Contratos Pendentes", pendentes, delta=f"{pendentes} verificar", delta_color="inverse")

            # Estilização
            def highlight_verificar(val):
                return 'color: red; font-weight: bold' if val == "Verificar" else ''

            st.dataframe(
                df_conferencia.style.applymap(highlight_verificar, subset=['Contrato CliqCCEE']),
                hide_index=True, use_container_width=True
            )

            # Download
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_conferencia.to_excel(writer, index=False, sheet_name='Book')
            st.download_button("📥 Baixar Excel", output.getvalue(), f"Book_{mes_nome_sel}.xlsx")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
