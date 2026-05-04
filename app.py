import streamlit as st
import pandas as pd
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES DE APOIO
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "":
        return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj))
    apenas_numeros = apenas_numeros.zfill(14)
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
    """Garante que a boleta seja string pura, sem .0 e sem espaços"""
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")
arquivo_subido = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'], key="atual")
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior.xlsx", type=['xlsx'], key="anterior")
arquivo_pessoas = st.sidebar.file_uploader("3. RelPers_858 (4).xlsx", type=['xlsx'], key="pessoas")

st.title("📑 Book de Energia")

# 4. PROCESSAMENTO DAS BASES DE APOIO (DICIONÁRIOS)
dict_mes_anterior = {}
dict_vendedor = {}
dict_comprador = {}

# Mês Anterior
if arquivo_anterior:
    try:
        df_apoio = pd.read_excel(arquivo_anterior)
        # Assume Col 0 = Boleta, Col 1 = Código Contrato
        df_apoio['chave'] = df_apoio.iloc[:, 0].apply(tratar_chave)
        dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio['chave'].values).to_dict()
        st.sidebar.success("✅ Mês Anterior carregado!")
    except Exception as e:
        st.sidebar.error(f"Erro no Mês Anterior: {e}")

# Relatório de Pessoas
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        # Col B (1) = Comprador | Col C (2) = Vendedor | Col D (3) = Boleta
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
        st.sidebar.success("✅ Relatório Pessoas carregado!")
    except Exception as e:
        st.sidebar.error(f"Erro no Relatório Pessoas: {e}")

# 5. PROCESSAMENTO DA BASE PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Identificação de colunas da base principal
        col_boleta = df_bruto.columns[0]
        col_operacao = df_bruto.columns[1]
        col_cnpj = df_bruto.columns[4]
        col_energia = df_bruto.columns[5]
        col_contraparte = df_bruto.columns[6]
        col_volume_mwh = df_bruto.columns[20]
        col_horas_mes = df_bruto.columns[15]
        col_mod_min = df_bruto.columns[28]
        col_mod_max = df_bruto.columns[29]
        col_cliq_para = df_bruto.columns[60]
        col_parte = df_bruto.columns[62]
        col_mod_wbc = df_bruto.columns[63]

        # Criar base de conferência
        df_conferencia = df_bruto[[col_boleta]].drop_duplicates()
        df_conferencia['Boleta_Formatada'] = df_conferencia[col_boleta].apply(tratar_chave)
        df_conferencia = df_conferencia.sort_values(by='Boleta_Formatada')

        # Helper para buscar dados na bruta
        df_lookup = df_bruto.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

        # Preenchimento
        df_conferencia['Operação'] = df_conferencia[col_boleta].map(df_lookup[col_operacao]).astype(str)
        
        trad_en = {"Incentivada-50%": "Incentivada-I5", "Incentivada-CQ50%": "Incentivada-CQ5", "Incentivada-100%": "Incentivada-I1", "Incentivada-0%": "Incentivada-I0", "Convencional": "Convencional"}
        df_conferencia['Tipo de Energia'] = df_conferencia[col_boleta].map(df_lookup[col_energia]).replace(trad_en)
        df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[col_parte]).astype(str)
        df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_contraparte])
        df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cnpj]).apply(formatar_cnpj)
        
        v_mwh = df_conferencia[col_boleta].map(df_lookup[col_volume_mwh])
        h_mes = df_conferencia[col_boleta].map(df_lookup[col_horas_mes])
        df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(4)
        
        df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[col_cliq_para])
        df_conferencia['Modulação WBC'] = df_conferencia[col_boleta].map(df_lookup[col_mod_wbc]).apply(limpar_modulacao)
        df_conferencia['Modulação Mínima'] = df_conferencia[col_boleta].map(df_lookup[col_mod_min])
        df_conferencia['Modulação Máxima'] = df_conferencia[col_boleta].map(df_lookup[col_mod_max])

        # --- BUSCAS COM CHAVE LIMPA ---
        df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta_Formatada'].map(dict_mes_anterior).fillna("-")
        df_conferencia['Comprador'] = df_conferencia['Boleta_Formatada'].map(dict_comprador).fillna("N/A")
        df_conferencia['Vendedor'] = df_conferencia['Boleta_Formatada'].map(dict_vendedor).fillna("N/A")

        # --- EXIBIÇÃO NA ORDEM SOLICITADA ---
        colunas_exibicao = [
            col_boleta, 'Operação', 'Tipo de Energia', 'Parte', 'Contraparte', 'CNPJ Contraparte', 
            'Volume MWm', 'CliqCCEE Paradigma', 'Modulação WBC', 'Modulação Mínima', 'Modulação Máxima', 
            'Contrato CliqCCEE mês anterior', # Penúltima seção
            'Comprador', 'Vendedor'           # Últimas colunas
        ]

        st.markdown("---")
        st.dataframe(
            df_conferencia[colunas_exibicao], 
            hide_index=True, 
            column_config={
                'Volume MWm': st.column_config.NumberColumn("Volume MWm", format="%.4f"),
            },
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Erro no processamento principal: {e}")
else:
    st.info("Por favor, faça o upload dos arquivos para começar.")
