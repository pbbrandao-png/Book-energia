import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES COM CACHE (Otimização de Performance)
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

def limpar_modulacao(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper()
    if "FLAT" in t: return "Flat"
    if "CARGA" in t: return "Carga"
    if "DECLARADO" in t or "INFORMADO" in t: return "Declarado"
    if "GERA" in t: return "Geracao"
    return texto

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

meses_nomes = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
# Define o índice 3 (Abril) como padrão conforme solicitado
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

# BOTÃO DE PROCESSAMENTO (Evita que o app rode a cada alteração de arquivo)
processar = st.sidebar.button("🚀 Processar Dados", use_container_width=True)

# 4. EXECUÇÃO
st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

if processar and arquivo_subido:
    with st.spinner('Processando bases... Isso pode levar alguns segundos.'):
        try:
            # Carregamento cacheado das bases CCEE
            db_matrix = carregar_csv_cliq_cached([arq_ccear, arq_cbr, arq_cceal1])
            db_bismut = carregar_csv_cliq_cached([arq_cceal2])

            # Leitura da base principal
            df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
            
            # Mapeamento de colunas por índice (mantendo a estrutura original)
            col_boleta = df_bruto.columns[0]
            col_operacao = df_bruto.columns[1]
            col_cnpj = df_bruto.columns[4]
            col_contraparte = df_bruto.columns[6]
            col_mes = df_bruto.columns[14]
            col_horas = df_bruto.columns[15]
            col_vol = df_bruto.columns[20]
            col_mod_min = df_bruto.columns[28]
            col_mod_max = df_bruto.columns[29]
            col_cliq_p = df_bruto.columns[60]
            col_parte = df_bruto.columns[62]
            col_mod_wbc = df_bruto.columns[63]

            df_bruto[col_mes] = pd.to_numeric(df_bruto[col_mes], errors='coerce')
            df_filtrada = df_bruto[df_bruto[col_mes] == mes_num_sel].copy()

            if df_filtrada.empty:
                st.warning(f"Nenhuma operação encontrada para o mês {mes_num_sel}.")
            else:
                # Dicionários de apoio (Pessoas e Mês Anterior)
                dict_mes_ant, dict_comp, dict_vend = {}, {}, {}
                
                if arquivo_anterior:
                    df_ant = pd.read_excel(arquivo_anterior, dtype=str)
                    dict_mes_ant = pd.Series(df_ant.iloc[:, 1].values, index=df_ant.iloc[:, 0].apply(tratar_chave).values).to_dict()

                if arquivo_pessoas:
                    df_pers = pd.read_excel(arquivo_pessoas)
                    df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
                    dict_comp = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
                    dict_vend = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()

                # Construção do DataFrame final com TODAS as colunas solicitadas anteriormente
                df_res = pd.DataFrame()
                df_res['Boleta'] = df_filtrada[col_boleta].apply(tratar_chave)
                df_res['Operacao'] = df_filtrada[col_operacao].astype(str)
                df_res['Parte'] = df_filtrada[col_parte].astype(str).str.strip()
                df_res['Contraparte'] = df_filtrada[col_contraparte]
                df_res['CNPJ Contraparte'] = df_filtrada[col_cnpj].apply(formatar_cnpj)
                df_res['Volume MWm'] = (df_filtrada[col_vol] / df_filtrada[col_horas]).fillna(0).round(4)
                df_res['CliqCCEE Paradigma'] = df_filtrada[col_cliq_p].apply(tratar_chave)
                df_res['Modulacao WBC'] = df_filtrada[col_mod_wbc].apply(limpar_modulacao)
                df_res['Modulacao Minima'] = df_filtrada[col_mod_min]
                df_res['Modulacao Maxima'] = df_filtrada[col_mod_max]
                df_res['Contrato CliqCCEE mes anterior'] = df_res['Boleta'].map(dict_mes_ant).fillna("-")
                df_res['Comprador'] = df_res['Boleta'].map(dict_comp).fillna("N/A")
                df_res['Vendedor'] = df_res['Boleta'].map(dict_vend).fillna("N/A")

                # Lógica de validação do contrato na CCEE
                def validar_cliq(row):
                    db = db_bismut if 'BISMUT' in str(row['Parte']).upper() else db_matrix
                    if db is None: return "Verificar"
                    
                    # Testa primeiro o paradigma, depois o mês anterior
                    for cod in [row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior']]:
                        c = tratar_chave(cod)
                        if c and c in db.index:
                            status = str(db.loc[c, 'SITUACAO_CONTRATO']).upper() if 'SITUACAO_CONTRATO' in db.columns else ""
                            if "RASCUNHO" not in status: return c
                    return "Verificar"

                df_res['Contrato CliqCCEE'] = df_res.apply(validar_cliq, axis=1)

                # --- EXIBIÇÃO ---
                pendentes = len(df_res[df_res['Contrato CliqCCEE'] == "Verificar"])
                st.metric("Contratos Pendentes", pendentes, delta=f"{pendentes} verificar", delta_color="inverse")

                # Estilização: destaque em vermelho para "Verificar"
                # Usando .map() para compatibilidade com versões novas do Pandas Styler
                def style_cliq(val):
                    return 'color: red; font-weight: bold' if val == "Verificar" else ''

                st.dataframe(
                    df_res.style.map(style_cliq, subset=['Contrato CliqCCEE']), 
                    hide_index=True, 
                    use_container_width=True
                )

                # Opção de Download
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_res.to_excel(writer, index=False, sheet_name='Conferencia')
                st.download_button("📥 Baixar Excel Completo", buffer.getvalue(), f"Book_Energia_{mes_nome_sel}.xlsx")

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
else:
    if not processar:
        st.info("Configure os arquivos na lateral e clique em 'Processar Dados' para gerar o relatório.")
