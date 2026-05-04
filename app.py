import streamlit as st
import pandas as pd
import re
import io

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES DE SUPORTE
@st.cache_data
def carregar_csv_cliq_cached(arquivo_lista):
    dfs = []
    for arquivo in arquivo_lista:
        if arquivo is not None:
            try:
                # Tenta ler garantindo que tudo venha como string e removendo espaços
                if arquivo.name.endswith('.csv'):
                    df = pd.read_csv(arquivo, sep='\t', encoding='latin-1', skiprows=1, dtype=str)
                else:
                    df = pd.read_excel(arquivo, dtype=str)
                
                if 'CODIGO_CONTRATO' in df.columns:
                    # Limpeza agressiva: remove .0, espaços e garante que é string
                    df['CODIGO_CONTRATO'] = df['CODIGO_CONTRATO'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
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
    if pd.isna(valor) or str(valor).strip().lower() in ['none', 'nan', '', '-']: return ""
    s = str(valor).strip()
    return re.sub(r'\.0$', '', s) # Remove o .0 do final se houver

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

processar = st.sidebar.button("🚀 Processar Dados", use_container_width=True)

# 4. EXECUÇÃO
st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

if processar and arquivo_subido:
    with st.spinner('Processando...'):
        try:
            db_matrix = carregar_csv_cliq_cached([arq_ccear, arq_cbr, arq_cceal1])
            db_bismut = carregar_csv_cliq_cached([arq_cceal2])

            df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
            
            # Mapeamento por índice
            indices = {
                'boleta': 0, 'operacao': 1, 'cnpj': 4, 'contraparte': 6,
                'mes': 14, 'horas': 15, 'vol': 20, 'mod_min': 28,
                'mod_max': 29, 'cliq_p': 60, 'parte': 62, 'mod_wbc': 63
            }

            df_bruto.iloc[:, indices['mes']] = pd.to_numeric(df_bruto.iloc[:, indices['mes']], errors='coerce')
            df_filtrada = df_bruto[df_bruto.iloc[:, indices['mes']] == mes_num_sel].copy()

            if df_filtrada.empty:
                st.warning(f"Sem dados para o mês {mes_num_sel}.")
            else:
                dict_mes_ant, dict_comp, dict_vend = {}, {}, {}
                
                if arquivo_anterior:
                    df_ant = pd.read_excel(arquivo_anterior, dtype=str)
                    dict_mes_ant = pd.Series(df_ant.iloc[:, 1].values, index=df_ant.iloc[:, 0].apply(tratar_chave).values).to_dict()

                if arquivo_pessoas:
                    df_pers = pd.read_excel(arquivo_pessoas)
                    df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
                    dict_comp = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
                    dict_vend = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()

                # Construção do DataFrame
                df_res = pd.DataFrame()
                df_res['Boleta'] = df_filtrada.iloc[:, indices['boleta']].apply(tratar_chave)
                df_res['Operacao'] = df_filtrada.iloc[:, indices['operacao']].astype(str)
                df_res['Parte'] = df_filtrada.iloc[:, indices['parte']].astype(str).str.strip()
                df_res['Contraparte'] = df_filtrada.iloc[:, indices['contraparte']]
                df_res['CNPJ Contraparte'] = df_filtrada.iloc[:, indices['cnpj']].apply(formatar_cnpj)
                df_res['Volume MWm'] = (df_filtrada.iloc[:, indices['vol']] / df_filtrada.iloc[:, indices['horas']]).fillna(0).round(4)
                df_res['CliqCCEE Paradigma'] = df_filtrada.iloc[:, indices['cliq_p']].apply(tratar_chave)
                df_res['Modulacao WBC'] = df_filtrada.iloc[:, indices['mod_wbc']].apply(limpar_modulacao)
                df_res['Modulacao Minima'] = df_filtrada.iloc[:, indices['mod_min']]
                df_res['Modulacao Maxima'] = df_filtrada.iloc[:, indices['mod_max']]
                df_res['Contrato CliqCCEE mes anterior'] = df_res['Boleta'].map(dict_mes_ant).fillna("-")
                df_res['Comprador'] = df_res['Boleta'].map(dict_comp).fillna("N/A")
                df_res['Vendedor'] = df_res['Boleta'].map(dict_vend).fillna("N/A")

                # Validação com busca aprimorada
                def validar_cliq(row):
                    db = db_bismut if 'BISMUT' in str(row['Parte']).upper() else db_matrix
                    if db is None: return "Verificar"
                    
                    # Tenta Paradigma, depois Mês Anterior
                    for cod in [row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior']]:
                        c = tratar_chave(cod)
                        if c and c in db.index:
                            status = str(db.loc[c, 'SITUACAO_CONTRATO']).upper() if 'SITUACAO_CONTRATO' in db.columns else ""
                            if "RASCUNHO" not in status: return c
                    return "Verificar"

                df_res['Contrato CliqCCEE'] = df_res.apply(validar_cliq, axis=1)

                # Ordenação decrescente
                df_res['Boleta_Num'] = pd.to_numeric(df_res['Boleta'], errors='coerce')
                df_res = df_res.sort_values(by='Boleta_Num', ascending=False).drop(columns=['Boleta_Num'])

                # Exibição
                pendentes = len(df_res[df_res['Contrato CliqCCEE'] == "Verificar"])
                st.metric("Contratos Pendentes", pendentes, delta=f"{pendentes} verificar", delta_color="inverse")

                st.dataframe(
                    df_res.style.map(lambda x: 'color: red; font-weight: bold' if x == "Verificar" else '', subset=['Contrato CliqCCEE']), 
                    hide_index=True, use_container_width=True
                )

                # Download usando o motor padrão (evita o erro do xlsxwriter)
                buffer = io.BytesIO()
                df_res.to_excel(buffer, index=False)
                st.download_button("📥 Baixar Excel", buffer.getvalue(), f"Book_{mes_nome_sel}.xlsx")

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
