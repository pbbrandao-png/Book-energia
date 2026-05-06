import streamlit as st
import pandas as pd
import numpy as np

# --- CONFIGURAÇÕES DA PÁGINA ---
st.set_page_config(page_title="Conferência de Contratos", layout="wide")

# --- FUNÇÕES DE UTILIDADE ---
def tratar_chave(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return "-"
    try:
        # Remove decimais se for numérico e converte para string limpa
        return str(int(float(valor))).strip()
    except:
        return str(valor).strip()

def formatar_cnpj(valor):
    if pd.isna(valor): return "-"
    cnpj = ''.join(filter(str.isdigit, str(valor)))
    if len(cnpj) == 14:
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return cnpj

def limpar_modulacao(valor):
    val = str(valor).strip().upper()
    return "Flat" if val in ["FLAT", "NAN", "NONE", ""] else val

def buscar_cliq_ccee(boleta_wbc, boleta_anterior, df_ccee, origem, vendedor, comprador):
    """
    Lógica de busca: 
    1. Tenta pela Boleta WBC do mês atual
    2. Tenta pela Boleta WBC do mês anterior
    3. Tenta pelo par Vendedor/Comprador
    """
    if df_ccee is None: return "Verificar"
    
    # 1. Busca por Boleta Atual
    if boleta_wbc != "-":
        match = df_ccee[df_ccee['Cod. Contrato'].astype(str) == boleta_wbc]
        if not match.empty: return match.iloc[0]['Cod. Contrato CCEE']
    
    # 2. Busca por Boleta Anterior
    if boleta_anterior != "-":
        match = df_ccee[df_ccee['Cod. Contrato'].astype(str) == boleta_anterior]
        if not match.empty: return match.iloc[0]['Cod. Contrato CCEE']

    # 3. Busca por Vendedor/Comprador (Lógica simplificada)
    if vendedor != "" and comprador != "":
        match = df_ccee[(df_ccee['Vendedor'].astype(str).str.contains(vendedor, case=False, na=False)) & 
                        (df_ccee['Comprador'].astype(str).str.contains(comprador, case=False, na=False))]
        if not match.empty: return match.iloc[0]['Cod. Contrato CCEE']

    return "Verificar"

# --- INICIALIZAÇÃO DO ESTADO ---
if 'df_bruto' not in st.session_state:
    st.session_state['df_bruto'] = None
    st.session_state['dict_mapa'] = {}
    st.session_state['dict_pendencias'] = {}
    st.session_state['dict_mes_anterior'] = {}
    st.session_state['dict_vendedor'] = {}
    st.session_state['dict_comprador'] = {}
    st.session_state['db_bismut'] = None
    st.session_state['db_ccear'] = None
    st.session_state['db_cbr'] = None
    st.session_state['db_matrix'] = None

# --- SIDEBAR: UPLOAD E FILTROS ---
with st.sidebar:
    st.header("📂 Upload de Arquivos")
    
    # 1. Base Principal (WBC)
    file_wbc = st.file_uploader("Contratos Aprovados (WBC)", type=['xlsm', 'xlsx'])
    
    # 2. ERP / Mapa de Boletas
    file_erp = st.file_uploader("Mapa ERP (Boleta vs Contrato)", type=['xlsx'])
    
    # 3. Pendências Financeiras
    file_pend = st.file_uploader("Pendências Financeiras", type=['xlsx'])
    
    # 4. Bases CCEE (Opcionais)
    with st.expander("Bases CCEE (Cliq)"):
        f_bismut = st.file_uploader("Newave Bismut", type=['xlsx'])
        f_matrix = st.file_uploader("Matrix", type=['xlsx'])
        f_cbr = st.file_uploader("CBR", type=['xlsx'])

    st.divider()
    mes_sel = st.selectbox("Mês de Referência", 
                          ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                           "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"], index=4)
    mes_num_sel = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", 
                   "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"].index(mes_sel) + 1

# --- PROCESSAMENTO DOS UPLOADS ---
if file_wbc and st.session_state['df_bruto'] is None:
    with st.spinner("Lendo base WBC..."):
        df = pd.read_excel(file_wbc, sheet_name='Contratos_Selecionados')
        st.session_state['df_bruto'] = df
        
        # Dicionários de apoio (Vendedor/Comprador/Mes Anterior)
        col_boleta = df.columns[0]
        st.session_state['dict_vendedor'] = df.set_index(col_boleta)[df.columns[66]].to_dict()
        st.session_state['dict_comprador'] = df.set_index(col_boleta)[df.columns[67]].to_dict()
        st.session_state['dict_mes_anterior'] = df.set_index(col_boleta)[df.columns[61]].apply(tratar_chave).to_dict()

if file_erp:
    df_erp = pd.read_excel(file_erp)
    # Supondo colunas: 'Boleta' e 'Contrato ERP'
    st.session_state['dict_mapa'] = df_erp.set_index(df_erp.columns[0])[df_erp.columns[1]].to_dict()

if file_pend:
    df_p = pd.read_excel(file_pend)
    # Coluna E (índice 4) = Cliente | Coluna I (índice 8) = Valor
    df_p_clean = df_p.iloc[:, [4, 8]].dropna()
    df_p_clean.columns = ['Cliente', 'Valor']
    st.session_state['dict_pendencias'] = df_p_clean.groupby('Cliente')['Valor'].sum().to_dict()

# Carregamento bases CCEE
if f_bismut: st.session_state['db_bismut'] = pd.read_excel(f_bismut)
if f_matrix: st.session_state['db_matrix'] = pd.read_excel(f_matrix)

# --- CORPO PRINCIPAL ---
st.title(f"📊 Conferência de Faturamento - {mes_sel}")

if st.session_state['df_bruto'] is not None:
    try:
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        
        # Filtro de Mês
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            # Lookup para agilizar busca
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Mapeamento de Colunas (WBC)
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]])
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            
            mapa_energia = {'Incentivada-50%': 'Incentivada-I5', 'Incentivada-100%': 'Incentivada-I1', 'Incentivada-0%': 'Incentivada-I0'}
            df_conferencia['Tipo Energia'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]]).astype(str).str.strip().replace(mapa_energia)

            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})
            
            # Cálculos de Volume
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0)

            # Modulações e Cliq
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            
            # --- COLUNAS AC E AD (DINÂMICAS PELO ÍNDICE) ---
            col_ac_nome = df_base.columns[28] # % Modulacao Min
            col_ad_nome = df_base.columns[29] # % Modulacao Max
            df_conferencia['Modulação Mínima'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_ac_nome]), errors='coerce').fillna(0)
            df_conferencia['Modulação Máxima'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_ad_nome]), errors='coerce').fillna(0)

            # Cruzamentos Externos
            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")

            def resolver_cliq(row):
                vend = str(row['Vendedor']) if row['Vendedor'] != "-" else ""
                comp = str(row['Comprador']) if row['Comprador'] != "-" else ""
                if 'BISMUT' in str(row['Parte']).upper():
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_matrix'], 'matrix', vend, comp)

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)
            
            # Pendência Financeira (Chave por Razão Social em maiúsculo)
            df_conferencia['Pendência Financeira'] = df_conferencia['Razao Social'].str.upper().map(st.session_state['dict_pendencias']).fillna(0.0)

            # --- EXIBIÇÃO ---
            ordem = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP', 
                'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 
                'CliqCCEE Paradigma', 'Modulacao WBC', 'Modulação Mínima', 'Modulação Máxima',
                'Vendedor', 'Comprador', 'Contrato CliqCCEE', 'Situacao ERP', 'Razao Social', 'Pendência Financeira'
            ]
            
            st.success(f"Foram encontrados {len(df_conferencia)} contratos para {mes_sel}.")
            
            st.dataframe(
                df_conferencia[ordem].sort_values(by=col_boleta),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"),
                    "Volume MWm": st.column_config.NumberColumn(format="%.6f"),
                    "Modulação Mínima": st.column_config.NumberColumn(format="%.2f%%"),
                    "Modulação Máxima": st.column_config.NumberColumn(format="%.2f%%"),
                    "Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f")
                }
            )

            # Botão de Download
            csv = df_conferencia[ordem].to_csv(index=False).encode('utf-8-sig')
            st.download_button("📥 Baixar Relatório (CSV)", csv, f"Conferencia_{mes_sel}.csv", "text/csv")
            
        else:
            st.warning(f"Nenhum dado encontrado para o mês de {mes_sel}.")
            
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
else:
    st.info("Aguardando upload da base WBC (Contratos Aprovados)...")
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
        if 'CODIGO_CONTRATO' in df.columns:
            df['CODIGO_CONTRATO'] = df['CODIGO_CONTRATO'].apply(tratar_chave)
            df = df.set_index('CODIGO_CONTRATO')
            return df
        return None
    except Exception: return None

# ─────────────────────────────────────────────────────────────────────────────
# MAPEAMENTO DE COLUNAS POR BASE CLIQ
# ─────────────────────────────────────────────────────────────────────────────
COLUNAS_CLIQ = {
    'matrix': {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'bismut': {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'cbr':    {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'ccear':  {'vendedor': 'SIGLA_PERFIL_VENDEDOR',  'comprador': 'SIGLA_PERFIL_COMPRADOR'},
}

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq, tipo_base, nome_vendedor, nome_comprador):
    if df_cliq is None: return "Verificar"
    mapa = COLUNAS_CLIQ.get(tipo_base, {})
    col_vend, col_comp = mapa.get('vendedor'), mapa.get('comprador')

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index: return False
        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame): row = row.iloc[0]
        if str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper() == 'RASCUNHO': return False
        if col_vend and col_vend in df_cliq.columns:
            if limpar_str(nome_vendedor) and limpar_str(row.get(col_vend, '')) != limpar_str(nome_vendedor): return False
        if col_comp and col_comp in df_cliq.columns:
            if limpar_str(nome_comprador) and limpar_str(row.get(col_comp, '')) != limpar_str(nome_comprador): return False
        return True

    if checar(cod_paradigma): return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior): return tratar_chave(cod_mes_anterior)
    return "Verificar"

# 3. INICIALIZAÇÃO DO SESSION STATE
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

# 4. INTERFACE LATERAL
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

# 5. CARREGAMENTO DOS DADOS (Com cache de ID)
if get_file_id(arquivo_subido) != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    if arquivo_subido:
        st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

# --- CORREÇÃO: LÓGICA DE SOMA DE PENDÊNCIA COM NORMALIZAÇÃO ---
if get_file_id(arquivo_pendencias) != st.session_state['fid_pendencias']:
    st.session_state['fid_pendencias'] = get_file_id(arquivo_pendencias)
    if arquivo_pendencias:
        try:
            df_p = pd.read_excel(arquivo_pendencias)
            df_p_simples = df_p.iloc[:, [4, 8]].copy()
            df_p_simples.columns = ['razao_social_pend', 'valor_pendente']
            df_p_simples['valor_pendente'] = pd.to_numeric(df_p_simples['valor_pendente'], errors='coerce').fillna(0)
            # CORREÇÃO: normalizar a chave para evitar falhas de match
            df_p_simples['razao_social_pend'] = (
                df_p_simples['razao_social_pend']
                .astype(str)
                .str.strip()
                .str.upper()
            )
            df_somado = df_p_simples.groupby('razao_social_pend')['valor_pendente'].sum().reset_index()
            st.session_state['dict_pendencias'] = dict(zip(df_somado['razao_social_pend'], df_somado['valor_pendente']))
        except Exception as e:
            st.session_state['dict_pendencias'] = {}
            st.warning(f"Erro ao carregar pendências: {e}")
# --------------------------------------------------------------

if get_file_id(arquivo_anterior) != st.session_state['fid_anterior']:
    st.session_state['fid_anterior'] = get_file_id(arquivo_anterior)
    if arquivo_anterior:
        df_apoio = pd.read_excel(arquivo_anterior, dtype=str)
        st.session_state['dict_mes_anterior'] = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].apply(tratar_chave).values).to_dict()

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

# 6. PROCESSAMENTO DA TABELA
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
            
            mapa_energia = {'Incentivada-50%': 'Incentivada-I5', 'Incentivada-100%': 'Incentivada-I1', 'Incentivada-0%': 'Incentivada-I0', 'Incentivada-CQ50%': 'Incentivada-CQ5'}
            df_conferencia['Tipo Energia'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]]).astype(str).str.strip().replace(mapa_energia)
            df_conferencia.loc[df_conferencia['Parte'].str.upper() == 'UFV JACARANDA 1', 'Tipo Energia'] = 'Incentivada-I5'

            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)
            
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)

            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            # Lógica CliqCCEE
            def resolver_cliq(row):
                vend, comp = (row['Vendedor'] if row['Vendedor'] != "-" else ""), (row['Comprador'] if row['Comprador'] != "-" else "")
                if 'BISMUT' in str(row['Parte']).upper(): 
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear','db_ccear'), ('cbr','db_cbr'), ('matrix','db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)

            # --- CORREÇÃO: MAPEAMENTO DA PENDÊNCIA COM NORMALIZAÇÃO ---
            df_conferencia['Pendência Financeira'] = (
                df_conferencia['Razao Social']
                .str.strip()
                .str.upper()
                .map(st.session_state['dict_pendencias'])
                .fillna(0.0)
            )
            # ----------------------------------------------------------

            # FILTROS
            st.write("### Filtros")
            f1, f2, f3, f4 = st.columns([2, 2, 2, 1])
            op_f = f1.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operacao'].unique()))
            parte_f = f2.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique()))
            cliq_f = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].unique()))
            zerar_intra = f4.toggle("Zerar Intraportfólio", value=False)

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if cliq_f != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]
            if zerar_intra:
                mask = df_final['Vendedor'].str.lower().str.strip() == df_final['Comprador'].str.lower().str.strip()
                df_final.loc[mask, ['Montante MWh', 'Volume MWm']] = 0.0

            # MÉTRICAS
            st.markdown("---")
            m1, m2, m3 = st.columns(3)
            m1.metric("Qtd. Operações Compra", len(df_final[df_final['Operacao'].str.upper().str.contains('COMPRA', na=False)]))
            m2.metric("Qtd. Operações Venda", len(df_final[df_final['Operacao'].str.upper().str.contains('VENDA', na=False)]))
            m3.metric("Total de Boletas na Tela", len(df_final))
            st.markdown("---")

            # EXIBIÇÃO
            ordem = [col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP', 
                    'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 
                    'CliqCCEE Paradigma', 'Modulacao WBC', 'Vendedor', 'Comprador',
                    'Contrato CliqCCEE', 'Situacao ERP', 'Razao Social', 'Pendência Financeira']
            
            st.dataframe(df_final[ordem].sort_values(by=col_boleta), use_container_width=True, hide_index=True,
                         column_config={
                             "Montante MWh": st.column_config.NumberColumn(format="%.3f"), 
                             "Volume MWm": st.column_config.NumberColumn(format="%.6f"),
                             "Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f")
                         })
        else: st.warning("Sem dados para este período.")
    except Exception as e: st.error(f"Erro: {e}")
