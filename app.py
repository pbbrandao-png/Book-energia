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
