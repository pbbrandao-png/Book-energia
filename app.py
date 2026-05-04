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

# Seletor de Período (Vigência)
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
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas = st.sidebar.file_uploader("3. RelPers_858 (4).xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_matrix = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx'])
arq_bismut = st.sidebar.file_uploader("Cliq Bismut", type=['xlsx'])
arq_cbr = st.sidebar.file_uploader("Cliq CBR", type=['xlsx'])
arq_lee = st.sidebar.file_uploader("Cliq LEE", type=['xlsx'])

st.title(f"📑 Book de Energia - {mes_nome_sel}/{ano_sel}")

# 4. PROCESSAMENTO DAS BASES DE APOIO
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

dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_apoio = pd.read_excel(arquivo_anterior)
        df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].apply(tratar_chave)
        dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()
    except: st.sidebar.error("Erro no arquivo Mês Anterior")

dict_vendedor, dict_comprador = {}, {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
    except: st.sidebar.error("Erro no arquivo Pessoas")

# 5. PROCESSAMENTO DA BASE PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de Colunas
        col_boleta = df_bruto.columns[0]
        col_operacao = df_bruto.columns[1]
        col_cnpj = df_bruto.columns[4]
        col_energia = df_bruto.columns[5]
        col_contraparte = df_bruto.columns[6]
        col_parte = df_bruto.columns[7]
        col_horas_mes = df_bruto.columns[15]
        col_volume_mwh = df_bruto.columns[20]
        col_mod_min = df_bruto.columns[28]
        col_mod_max = df_bruto.columns[29]
        col_cliq_para = df_bruto.columns[60]
        col_mod_wbc = df_bruto.columns[63]
        col_mes_suprimento = df_bruto.columns[14] # Coluna O

        # Filtro por Mês (Coluna O)
        df_bruto[col_mes_suprimento] = pd.to_numeric(df_bruto[col_mes_suprimento], errors='coerce')
        df_filtrada = df_bruto[df_bruto[col_mes_suprimento] == mes_num_sel].copy()

        if df_filtrada.empty:
            st.warning("Nenhuma operação encontrada para este mês na coluna O.")
        else:
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Preenchimento das colunas que tínhamos antes
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

            # Colunas de Apoio
            df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("N/A")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("N/A")

            # Lógica Contrato Cliq CCEE (Nova)
            def buscar_cliq_ccee(row):
                boleta = row['Boleta_Key']
                orig = df_lookup.loc[row[col_boleta]]
                validacao_local = f"{orig.iloc[19]}{orig.iloc[20]}{vigencia_match_ccee}"
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

            # Filtros Adicionais na Tela
            st.write("### Filtros Rápidos")
            f1, f2, f3 = st.columns(3)
            with f1: op_f = st.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operação'].unique().tolist()))
            with f2: parte_f = st.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique().tolist()))
            with f3: rem_zero = st.checkbox("Ocultar Zerados", value=False)

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operação'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if rem_zero: df_final = df_final[df_final['Volume MWm'] != 0]

            # Métricas
            m1, m2, m3 = st.columns(3)
            compra = df_final[df_final['Operação'].str.contains('Compra', case=False)]['Volume MWm'].sum()
            venda = df_final[df_final['Operação'].str.contains('Venda', case=False)]['Volume MWm'].sum()
            m1.metric("Boletas", len(df_final))
            m2.metric("Total Compra", f"{compra:.4f}")
            m3.metric("Total Venda", f"{venda:.4f}")

            # Exibição Final
            col_ordem = [
                col_boleta, 'Operação', 'Tipo de Energia', 'Parte', 'Contraparte', 'CNPJ Contraparte', 
                'Volume MWm', 'CliqCCEE Paradigma', 'Modulação WBC', 'Modulação Mínima', 'Modulação Máxima', 
                'Contrato Cliq CCEE', 'Contrato CliqCCEE mês anterior', 'Comprador', 'Vendedor'
            ]
            st.dataframe(df_final[col_ordem], hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Erro: {e}")
