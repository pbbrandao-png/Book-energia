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
    if "GERA" in t: return "Geração"
    return texto

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

# Seletor de Período
st.sidebar.subheader("📅 Período de Referência")
meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1 
anos = [str(a) for a in range(2024, 2031)]
ano_sel = st.sidebar.selectbox("Ano", anos, index=2) # 2026

vigencia_match_ccee = f"{str(mes_num_sel).zfill(2)}/{ano_sel}"

st.sidebar.markdown("---")
arquivo_subido = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior_3.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. RelPers_858 (4).xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE (CSVs Tabulados)")
arq_matrix = st.sidebar.file_uploader("Cliq Matrix", type=['csv'])
arq_bismut = st.sidebar.file_uploader("Cliq Bismut", type=['csv'])
arq_cbr = st.sidebar.file_uploader("Cliq CBR", type=['csv'])
arq_lee = st.sidebar.file_uploader("Cliq LEE", type=['csv'])

st.title(f"📑 Book de Energia - {mes_nome_sel}/{ano_sel}")

# 4. PROCESSAMENTO DAS BASES DE APOIO (Ajustado para CSVs da CCEE)
def carregar_cliq_csv(arquivo):
    if arquivo:
        try:
            # CSVs da CCEE usam tabulação (\t) e costumam ter 'sep=\t' na primeira linha
            df = pd.read_csv(arquivo, sep='\t', encoding='latin-1', skiprows=1)
            
            # Remove colunas fantasmas e garante que a coluna de Boleta (index 0 ou CODIGO_CONTRATO) seja tratada
            # Nas suas bases, CODIGO_CONTRATO é a primeira coluna real após o skip
            df['chave_boleta'] = df.iloc[:, 0].apply(tratar_chave)
            return df.set_index('chave_boleta')
        except Exception as e:
            st.sidebar.error(f"Erro ao ler CSV: {e}")
            return None
    return None

db_matrix = carregar_cliq_csv(arq_matrix)
db_bismut = carregar_cliq_csv(arq_bismut)
db_cbr = carregar_cliq_csv(arq_cbr)
db_lee = carregar_cliq_csv(arq_lee)

dict_mes_anterior = {}
if arquivo_anterior:
    try:
        # Lê Mês Anterior_3.xlsx
        df_apoio = pd.read_excel(arquivo_anterior)
        # Assume Boleta na Coluna A (0) e ID Cliq na Coluna B (1)
        df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].apply(tratar_chave)
        dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()
    except: st.sidebar.error("Erro ao ler Mês Anterior_3.xlsx")

dict_vendedor, dict_comprador = {}, {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
    except: pass

# 5. PROCESSAMENTO DA BASE PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de Colunas (Coluna BK é índice 62)
        col_boleta = df_bruto.columns[0]
        col_operacao = df_bruto.columns[1]
        col_cnpj = df_bruto.columns[4]
        col_energia = df_bruto.columns[5]
        col_contraparte = df_bruto.columns[6]
        col_parte_bk = df_bruto.columns[62] # Coluna BK
        col_mes_suprimento = df_bruto.columns[14] # Coluna O
        col_horas_mes = df_bruto.columns[15]
        col_volume_mwh = df_bruto.columns[20]
        col_mod_min = df_bruto.columns[28]
        col_mod_max = df_bruto.columns[29]
        col_cliq_para = df_bruto.columns[60]
        col_mod_wbc = df_bruto.columns[63]

        # Filtro de Mês
        df_bruto[col_mes_suprimento] = pd.to_numeric(df_bruto[col_mes_suprimento], errors='coerce')
        df_filtrada = df_bruto[df_bruto[col_mes_suprimento] == mes_num_sel].copy()

        if df_filtrada.empty:
            st.warning(f"Nenhuma operação para o mês {mes_num_sel}.")
        else:
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Atribuição de Colunas
            df_conferencia['Operação'] = df_conferencia[col_boleta].map(df_lookup[col_operacao]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[col_parte_bk]).astype(str).str.strip()
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_contraparte])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cnpj]).apply(formatar_cnpj)
            
            # Volume
            v_mwh = df_conferencia[col_boleta].map(df_lookup[col_volume_mwh])
            h_mes = df_conferencia[col_boleta].map(df_lookup[col_horas_mes])
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(4)
            
            # Modulações e Cliq
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[col_cliq_para])
            df_conferencia['Modulação WBC'] = df_conferencia[col_boleta].map(df_lookup[col_mod_wbc]).apply(limpar_modulacao)
            df_conferencia['Modulação Mínima'] = df_conferencia[col_boleta].map(df_lookup[col_lookup[col_mod_min]])
            df_conferencia['Modulação Máxima'] = df_conferencia[col_boleta].map(df_lookup[col_lookup[col_mod_max]])

            # Mês Anterior
            df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("N/A")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("N/A")

            # Busca ID Cliq CCEE Atual
            def buscar_cliq_ccee(row):
                boleta = row['Boleta_Key']
                orig = df_lookup.loc[row[col_boleta]]
                # Verificação customizada (Tipo + Submercado + Vigência)
                # No CSV, os índices mudam levemente. Adaptamos para buscar pela boleta:
                parte_val = str(orig.iloc[62]).upper()
                bases = [db_bismut] if "BISMUT" in parte_val else [db_matrix, db_cbr, db_lee]
                
                for db in bases:
                    if db is not None and boleta in db.index:
                        return boleta # Retorna a boleta se encontrada na base validada
                return "Verificar"

            df_conferencia['Contrato Cliq CCEE'] = df_conferencia.apply(buscar_cliq_ccee, axis=1)

            # Renderização Final
            ordem = [
                col_boleta, 'Operação', 'Parte', 'Contraparte', 'CNPJ Contraparte', 
                'Volume MWm', 'CliqCCEE Paradigma', 'Modulação WBC', 'Modulação Mínima', 'Modulação Máxima', 
                'Contrato CliqCCEE mês anterior', 'Comprador', 'Vendedor', 'Contrato Cliq CCEE'
            ]
            st.dataframe(df_conferencia[ordem], hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Erro: {e}")
