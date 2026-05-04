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
        
        # Ajuste o nome da coluna de índice conforme o cabeçalho real do seu CSV
        col_id = 'CODIGO_CONTRATO' 
        if col_id in df.columns:
            df[col_id] = df[col_id].apply(tratar_chave)
            df = df.set_index(col_id)
        return df
    except Exception:
        return None

def buscar_cliq_complexo(row, db_matrix, db_bismut):
    """
    Replica a lógica da fórmula: 
    - Escolhe a base (Bismut vs Matrix/CBR/LEE)
    - Tenta Cod_Principal e Cod_Alt
    - Valida se Status != RASCUNHO
    - Valida se Chave Concatenada (T+U+L) bate com a SIGLA_AGENTE na base
    """
    parte = str(row.get('Parte', '')).strip().upper()
    chave_validacao = str(row.get('Chave_Concatenada', '')).strip().upper()
    
    # Define o banco de dados alvo
    # Se na 'Parte' houver BISMUT, usa a base bismut, senão usa a matrix
    db_alvo = db_bismut if "BISMUT" in parte else db_matrix
    
    if db_alvo is None: return "-"

    # Tenta os dois códigos (S12 e O12 da sua fórmula)
    codigos_para_testar = [row.get('CliqCCEE Paradigma'), row.get('_cliq_alt')]
    
    for cod in codigos_para_testar:
        cod_str = tratar_chave(cod)
        if cod_str and cod_str in db_alvo.index:
            dados_cliq = db_alvo.loc[cod_str]
            if isinstance(dados_cliq, pd.DataFrame):
                dados_cliq = dados_cliq.iloc[0]
            
            # Critérios da fórmula:
            status = str(dados_cliq.get('SITUACAO_CONTRATO', '')).strip().upper()
            # No Excel você concatena T+U+L e compara com a coluna C da CLIQ
            agente_base = str(dados_cliq.get('SIGLA_AGENTE_COMPRADOR', '')).strip().upper()
            
            if status != "RASCUNHO" and agente_base == chave_validacao:
                return cod_str
                
    return "Verificar"

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

st.sidebar.subheader("Período de Referência")
meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1
ano_sel = st.sidebar.selectbox("Ano", [str(a) for a in range(2024, 2031)], index=2)

st.sidebar.markdown("---")
arquivo_subido   = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior (Opcional)", type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. RelPers_858 (Pessoas)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE - Matrix/CBR/LEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q_101457", type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado_101457", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq CCEAL Firme_101457", type=['xlsx', 'csv'])

st.sidebar.subheader("Bases Cliq CCEE - Bismut")
arq_cceal2 = st.sidebar.file_uploader("Cliq CCEAL Firme_101475 (Bismut)", type=['xlsx', 'csv'])

# 4. CARREGAMENTO E CONSOLIDAÇÃO DAS BASES CLIQ
# Matrix / CBR / LEE
dfs_matrix = []
for arq in [arq_ccear, arq_cbr, arq_cceal1]:
    df_temp = carregar_csv_cliq(arq)
    if df_temp is not None: dfs_matrix.append(df_temp)
db_matrix = pd.concat(dfs_matrix) if dfs_matrix else None

# Bismut
db_bismut = carregar_csv_cliq(arq_cceal2)

# Dicionários Auxiliares
dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_ant = pd.read_excel(arquivo_anterior, dtype=str)
        dict_mes_anterior = pd.Series(df_ant.iloc[:, 1].values, index=df_ant.iloc[:, 0].apply(tratar_chave).values).to_dict()
    except: pass

dict_vendedor, dict_comprador = {}, {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor  = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
    except: pass

# 5. PROCESSAMENTO PRINCIPAL
st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento Dinâmico de Colunas (ajuste os índices se a planilha mudar)
        # Note: Usei os nomes que você forneceu no código anterior
        col_boleta = df_bruto.columns[0]
        df_bruto[df_bruto.columns[14]] = pd.to_numeric(df_bruto[df_bruto.columns[14]], errors='coerce')
        df_filtrada = df_bruto[df_bruto[df_bruto.columns[14]] == mes_num_sel].copy()

        if not df_filtrada.empty:
            df_conferencia = pd.DataFrame()
            df_conferencia['Boleta'] = df_filtrada[col_boleta].apply(tratar_chave)
            
            # Campos Básicos
            df_conferencia['Operacao'] = df_filtrada[df_bruto.columns[1]].astype(str)
            df_conferencia['Parte'] = df_filtrada[df_bruto.columns[62]].astype(str).str.strip()
            df_conferencia['Contraparte'] = df_filtrada[df_bruto.columns[6]]
            df_conferencia['CNPJ Contraparte'] = df_filtrada[df_bruto.columns[4]].apply(formatar_cnpj)
            
            # Chave Concatenada (T + U + L do Excel)
            # Substitua pelos nomes ou índices reais das colunas T, U e L
            col_T = df_bruto.columns[19] # Exemplo
            col_U = df_bruto.columns[20] # Exemplo
            col_L = df_bruto.columns[11] # Exemplo
            df_conferencia['Chave_Concatenada'] = (
                df_filtrada[col_T].astype(str) + 
                df_filtrada[col_U].astype(str) + 
                df_filtrada[col_L].astype(str)
            )

            # IDs de busca para o Cliq
            df_conferencia['CliqCCEE Paradigma'] = df_filtrada[df_bruto.columns[60]].apply(tratar_chave)
            df_conferencia['_cliq_alt'] = df_filtrada[df_bruto.columns[14]].apply(tratar_chave)

            # Aplicação da Lógica Complexa (Substitui a fórmula da Coluna V)
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(
                lambda r: buscar_cliq_complexo(r, db_matrix, db_bismut), axis=1
            )

            # Exibição
            st.write("### Resultado da Conferência")
            st.dataframe(df_conferencia.drop(columns=['Chave_Concatenada', '_cliq_alt']), use_container_width=True)
            
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
