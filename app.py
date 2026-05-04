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

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def limpar_modulacao(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper()
    if "FLAT" in t: return "Flat"
    if "CARGA" in t: return "Carga"
    if "DECLARADO" in t or "INFORMADO" in t: return "Declarado"
    if "GERA" in t: return "Geração"
    return texto

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

# --- NOVO: SELETOR DE MÊS E ANO (Vigência) ---
st.sidebar.subheader("📅 Período de Referência")
meses = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]
anos = [str(a) for a in range(2024, 2031)]

col_mes, col_ano = st.sidebar.columns(2)
mes_sel = col_mes.selectbox("Mês", meses, index=datetime.now().month - 1)
ano_sel = col_ano.selectbox("Ano", anos, index=0)

vigencia_referencia = f"{mes_sel}/{ano_sel}" # Formato MM/AAAA para o match

# Uploads
st.sidebar.markdown("---")
arquivo_subido = st.sidebar.file_uploader("1. Base do Mês Atual", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. RelPers_858 (4).xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_matrix = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx'])
arq_bismut = st.sidebar.file_uploader("Cliq Bismut", type=['xlsx'])
arq_cbr = st.sidebar.file_uploader("Cliq CBR", type=['xlsx'])
arq_lee = st.sidebar.file_uploader("Cliq LEE", type=['xlsx'])

st.title(f"📑 Book de Energia - {vigencia_referencia}")

# 4. CARREGAMENTO DAS BASES CCEE
def carregar_cliq(arquivo):
    if arquivo:
        try:
            df = pd.read_excel(arquivo)
            df['chave_boleta'] = df.iloc[:, 3].apply(tratar_chave)
            return df.set_index('chave_boleta')
        except: return None
    return None

db_matrix = carregar_cliq(arq_matrix)
db_bismut = carregar_cliq(arq_bismut)
db_cbr = carregar_cliq(arq_cbr)
db_lee = carregar_cliq(arq_lee)

# 5. PROCESSAMENTO PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de colunas
        col_boleta = df_bruto.columns[0]
        col_vigencia_original = df_bruto.columns[11] # Col L
        
        # FILTRO INICIAL: Só processa o que for do Mês/Ano selecionado
        # Nota: Ajuste a lógica de conversão da Col L se ela não for string/data padrão
        df_bruto['vigencia_aux'] = df_bruto[col_vigencia_original].astype(str).str.strip()
        
        # Filtramos a base bruta antes de tudo
        df_filtrada_periodo = df_bruto[df_bruto['vigencia_aux'].str.contains(vigencia_referencia, na=False)].copy()

        if df_filtrada_periodo.empty:
            st.warning(f"Nenhuma operação encontrada para {vigencia_referencia} na base subida.")
        else:
            df_conferencia = df_filtrada_periodo[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            df_lookup = df_filtrada_periodo.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Lógica de Busca CCEE
            def buscar_cliq_ccee(row):
                boleta = row['Boleta_Key']
                orig = df_lookup.loc[row[col_boleta]]
                
                # Monta validação T + U + L (L agora vem do seletor fixo)
                t = str(orig.iloc[19]).strip() 
                u = str(orig.iloc[20]).strip()
                l = vigencia_referencia 
                validacao_local = f"{t}{u}{l}"
                
                parte = str(orig.iloc[7]).upper()
                bases = [db_bismut] if "BISMUT" in parte else [db_matrix, db_cbr, db_lee]
                
                for db in bases:
                    if db is not None and boleta in db.index:
                        info = db.loc[boleta]
                        if isinstance(info, pd.DataFrame): info = info.iloc[0]
                        if str(info.iloc[2]).strip() == validacao_local and str(info.iloc[10]) != "Rascunho":
                            return boleta
                return "Verificar"

            df_conferencia['Contrato Cliq CCEE'] = df_conferencia.apply(buscar_cliq_ccee, axis=1)
            
            # (Adicione aqui as outras colunas como Operação, Tipo de Energia, etc., mapeando da df_lookup)
            df_conferencia['Operação'] = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[1]])

            # Exibição
            colunas_exibicao = [col_boleta, 'Operação', 'Contrato Cliq CCEE'] # Adicione as demais aqui
            st.dataframe(df_conferencia[colunas_exibicao], hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Erro: {e}")
