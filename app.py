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

st.sidebar.subheader("📅 Período de Referência")
meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
# Converte nome para número (Janeiro = 1, Abril = 4)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1 

anos = [str(a) for a in range(2024, 2031)]
ano_sel = st.sidebar.selectbox("Ano", anos, index=2) # Index 2 costuma ser 2026 na lista

# Formato para o Match no CliqCCEE (T+U+L)
vigencia_match_ccee = f"{str(mes_num_sel).zfill(2)}/{ano_sel}" 

# Uploads
st.sidebar.markdown("---")
arquivo_subido = st.sidebar.file_uploader("1. Base do Mês Atual", type=['xlsx', 'xlsm'])
# ... (outros uploaders aqui)

st.title(f"📑 Book de Energia - {mes_nome_sel}/{ano_sel}")

# 4. CARREGAMENTO DAS BASES CCEE (Mantendo a lógica anterior)
def carregar_cliq(arquivo):
    if arquivo:
        try:
            df = pd.read_excel(arquivo)
            df['chave_boleta'] = df.iloc[:, 3].apply(tratar_chave)
            return df.set_index('chave_boleta')
        except: return None
    return None

# Supondo que você carregou as bases arq_matrix, etc.
# db_matrix = carregar_cliq(arq_matrix) ...

# 5. PROCESSAMENTO PRINCIPAL
if arquivo_subido:
    try:
        # Carregando a aba específica
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de Colunas
        col_boleta = df_bruto.columns[0]
        col_mes_suprimento = df_bruto.columns[14] # Coluna O (índice 14)
        
        # --- FILTRO POR MÊS (COLUNA O) ---
        # Convertemos a coluna O para numérico para não ter erro de comparação
        df_bruto[col_mes_suprimento] = pd.to_numeric(df_bruto[col_mes_suprimento], errors='coerce')
        
        # Filtramos apenas o mês selecionado no seletor
        df_filtrada_periodo = df_bruto[df_bruto[col_mes_suprimento] == mes_num_sel].copy()

        if df_filtrada_periodo.empty:
            st.warning(f"Nenhuma operação encontrada para o mês {mes_num_sel} na coluna O.")
        else:
            df_conferencia = df_filtrada_periodo[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            # Helper para busca rápida
            df_lookup = df_filtrada_periodo.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Lógica de Busca CCEE (Match T + U + L)
            def buscar_cliq_ccee(row):
                boleta = row['Boleta_Key']
                orig = df_lookup.loc[row[col_boleta]]
                
                # Dados para o match
                t = str(orig.iloc[19]).strip() # Col T
                u = str(orig.iloc[20]).strip() # Col U
                l = vigencia_match_ccee        # Usa o formato "04/2026"
                
                validacao_local = f"{t}{u}{l}"
                
                parte = str(orig.iloc[7]).upper()
                # (Lógica de seleção de bases: db_bismut ou [db_matrix, db_cbr, db_lee])
                # ... busca nas bases como fizemos antes ...
                return "Verificar" # Exemplo simplificado

            df_conferencia['Contrato Cliq CCEE'] = df_conferencia.apply(buscar_cliq_ccee, axis=1)

            # --- EXIBIÇÃO ---
            st.dataframe(df_conferencia, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
