import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# ─────────────────────────────────────────────────────────────────────────────
# LISTA DE DESTAQUE - CONTRATOS ESPECIAIS
# ─────────────────────────────────────────────────────────────────────────────
CONTRATOS_ESPECIAIS_CCEAR = [
    "2813298", "2813299", "2813300", "2813301", "2813302", "2813303", 
    "2813304", "2813305", "4159778", "4159779", "4159780", "4686267", 
    "4686268", "4686269", "4686270"
]

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

def limpar_str(valor):
    if pd.isna(valor) or valor == "": return ""
    return str(valor).strip().lower()

def get_file_id(arq): 
    return (arq.name, arq.size) if arq else None

def carregar_csv_cliq(arquivo):
    if arquivo is None: return None
    try:
        nome = arquivo.name if hasattr(arquivo, 'name') else str(arquivo)
        if nome.endswith('.csv'):
            df = pd.read_csv(arquivo, sep='\t', encoding='latin-1', skiprows=1, dtype=str)
        else:
            df = pd.read_excel(arquivo, dtype=str)
        return df # Retorna o DF completo para busca por colunas, não apenas índice
    except Exception: return None

# 3. NOVA LÓGICA DE CONFERÊNCIA (MATCH TRIPLO)
def buscar_por_match_triplo(row, df_cliq):
    if df_cliq is None or df_cliq.empty: return "Verificar"
    
    vend_alvo = limpar_str(row['Vendedor'])
    comp_alvo = limpar_str(row['Comprador'])
    sub_wbc = row['Submercado']
    
    # Conversão de submercado para padrão CCEE
    de_para_sub = {'Sudeste': 'se/co', 'Sul': 's', 'Norte': 'n', 'Nordeste': 'ne'}
    sub_alvo = de_para_sub.get(sub_wbc, sub_wbc).lower()

    # Filtros na base CCEE:
    # 1. Não pode ser Rascunho
    # 2. Match Vendedor + Comprador + Submercado
    mask = (
        (df_cliq['SITUACAO_CONTRATO'].str.upper() != 'RASCUNHO') &
        (df_cliq['SIGLA_PERFIL_VENDEDOR'].str.lower().str.strip() == vend_alvo) &
        (df_cliq['SIGLA_PERFIL_COMPRADOR'].str.lower().str.strip() == comp_alvo) &
        (df_cliq['SUBMERCADO_ENTREGA'].str.lower().str.strip() == sub_alvo)
    )
    
    df_match = df_cliq[mask]
    
    if not df_match.empty:
        # Retorna o código do contrato encontrado
        return tratar_chave(df_match.iloc[0]['CODIGO_CONTRATO'])
    
    return "Verificar"

# 4. INICIALIZAÇÃO DO SESSION STATE
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias',
              'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
    if chave not in st.session_state: st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state: st.session_state[chave] = None

if 'mes_sel' not in st.session_state: st.session_state['mes_sel'] = meses_nomes[datetime.now().month - 1]
if 'ano_sel' not in st.session_state: st.session_state['ano_sel'] = str(datetime.now().year)

# 5. INTERFACE LATERAL
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=meses_nomes.index(st.session_state['mes_sel']), key='mes_sel')
ano_sel_val = st.sidebar.selectbox("Ano", anos, index=anos.index(st.session_state['ano_sel']) if st.session_state['ano_sel'] in anos else 0, key='ano_sel')
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.markdown("---")
arquivo_subido    = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior  = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx",      type=['xlsx'])
arquivo_pessoas   = st.sidebar.file_uploader("3. Exportador (4).xlsx",          type=['xlsx'])
arquivo_mapa      = st.sidebar.file_uploader("4. Mapa Financeiro (Excel)",     type=['xlsx'])
arquivo_pendencias = st.sidebar.file_uploader("5. Pendências Financeiras (Excel)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear, arq_cbr = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv']), st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1, arq_cceal2 = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx', 'csv']), st.sidebar.file_uploader("Cliq Bismut", type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel_val}")

# 6. CARREGAMENTO DOS DADOS (Otimizado com IDs)
if get_file_id(arquivo_subido) != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    if arquivo_subido: st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

if get_file_id(arquivo_anterior) != st.session_state['fid_anterior']:
    st.session_state['fid_anterior'] = get_file_id(arquivo_anterior)
    if arquivo_anterior:
        try:
            df_apoio = pd.read_excel(arquivo_anterior, dtype=str)
            st.session_state['dict_mes_anterior'] = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].apply(tratar_chave).values).to_dict()
        except: st.session_state['dict_mes_anterior'] = {}

if get_file_id(arquivo_pendencias) != st.session_state['fid_pendencias']:
    st.session_state['fid_pendencias'] = get_file_id(arquivo_pendencias)
    if arquivo_pendencias:
        try:
            df_p = pd.read_excel(arquivo_pendencias)
            df_p_simples = df_p.iloc[:, [4, 8]].copy()
            df_p_simples.columns = ['razao_social_pend', 'valor_pendente']
            df_p_simples['valor_pendente'] = pd.to_numeric(df_p_simples['valor_pendente'], errors='coerce').fillna(0)
            df_p_simples['razao_social_pend'] = df_p_simples['razao_social_pend'].astype(str).str.strip().str.upper()
            df_somado = df_p_simples.groupby('razao_social_pend')['valor_pendente'].sum().reset_index()
            st.session_state['dict_pendencias'] = dict(zip(df_somado['razao_social_pend'], df_somado['valor_pendente']))
        except: st.session_state['dict_pendencias'] = {}

if get_file_id(arquivo_pessoas) != st.session_state['fid_pessoas']:
    st.session_state['fid_pessoas'] = get_file_id(arquivo_pessoas)
    if arquivo_pessoas:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        st.session_state['dict_comprador'] = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        st.session_state['dict_vendedor'] = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()

if get_file_id(arquivo_mapa) != st.session_state['fid_mapa']:
    st.session_state['fid_mapa'] = get_file_id(arquivo_mapa)
    if arquivo_mapa:
        df_m = pd.read_excel(arquivo_mapa)
        st.session_state['dict_mapa'] = pd.Series(df_m['Situacao_ERP'].values, index=df_m['Codigo_WBC'].apply(tratar_chave).values).to_dict()

if (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)) != st.session_state['chave_matrix']:
    st.session_state['chave_matrix'] = (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1))
    st.session_state['db_ccear'], st.session_state['db_cbr'], st.session_state['db_matrix'] = carregar_csv_cliq(arq_ccear), carregar_csv_cliq(arq_cbr), carregar_csv_cliq(arq_cceal1)

if get_file_id(arq_cceal2) != st.session_state['fid_cceal2']:
    st.session_state['fid_cceal2'] = get_file_id(arq_cceal2)
    st.session_state['db_bismut'] = carregar_csv_cliq(arq_cceal2)

# 7. PROCESSAMENTO DA TABELA
if st.session_state['df_bruto'] is not None:
    try:
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Colunas da Tabela
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            
            # --- EXECUÇÃO DA CONFERÊNCIA DE MATCH ---
            def resolver_cliq_final(row):
                if 'BISMUT' in str(row['Parte']).upper():
                    return buscar_por_match_triplo(row, st.session_state['db_bismut'])
                
                # Procura sequencial nas outras bases
                for base_key in ['db_matrix', 'db_ccear', 'db_cbr']:
                    res = buscar_por_match_triplo(row, st.session_state[base_key])
                    if res != "Verificar": return res
                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq_final, axis=1)

            # --- BALÕES DE MÉTRICAS (CONFERÊNCIA 1) ---
            boletas_sem_contrato = df_conferencia[df_conferencia['Contrato CliqCCEE'] == 'Verificar']
            qtd_erro = len(boletas_sem_contrato)
            
            b1, b2 = st.columns(2)
            with b1:
                st.metric("Total de Boletas", len(df_conferencia))
            with b2:
                label_status = "⚠️ Boletas Sem Contrato" if qtd_erro > 0 else "✅ Contratos OK"
                st.metric(label_status, qtd_erro, delta=f"{qtd_erro} pendentes" if qtd_erro > 0 else "Tudo certo", delta_color="inverse")

            if qtd_erro > 0:
                with st.expander("Clique para ver boletas sem match na CCEE", expanded=False):
                    st.dataframe(boletas_sem_contrato[[col_boleta, 'Parte', 'Vendedor', 'Comprador', 'Submercado']], use_container_width=True, hide_index=True)

            # --- RESTANTE DO PROCESSAMENTO (VBC, MWm, ETC) ---
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)
            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            
            # Exibição Principal
            st.write("---")
            st.write("### Detalhamento Geral")
            ordem_final = [col_boleta, 'Operacao', 'Parte', 'Contraparte', 'Submercado', 'Vendedor', 'Comprador', 'Contrato CliqCCEE', 'Volume MWm', 'Situacao ERP']
            st.dataframe(df_conferencia[ordem_final].sort_values(by=col_boleta), use_container_width=True, hide_index=True)

        else: st.warning("Sem dados para este período.")
    except Exception as e: st.error(f"Erro: {e}")
