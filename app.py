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

# 5. CARREGAMENTO DOS DADOS
if get_file_id(arquivo_subido) != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    if arquivo_subido:
        st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

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
            df_conferencia['Tipo Energia'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]]).astype(str).str.strip()
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]])
            
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)

            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            df_conferencia['% Modulacao Min'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[28]]), errors='coerce')
            df_conferencia['% Modulacao Max'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[29]]), errors='coerce')
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            def resolver_cliq(row):
                vend, comp = (row['Vendedor'] if row['Vendedor'] != "-" else ""), (row['Comprador'] if row['Comprador'] != "-" else "")
                if 'BISMUT' in str(row['Parte']).upper(): 
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear','db_ccear'), ('cbr','db_cbr'), ('matrix','db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)

            # LÓGICA DE SOMA (Volume BOOK)
            df_soma_cliq = df_conferencia[~df_conferencia['Contrato CliqCCEE'].isin(['Verificar', '-', ''])].copy()
            dict_soma_book = df_soma_cliq.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['Volume BOOK'] = df_conferencia['Contrato CliqCCEE'].map(dict_soma_book).fillna(0.0).round(6)

            # ─────────────────────────────────────────────────────────────────────────────
            # NOVA LÓGICA: Volume CliqCCEE (Busca MWmedio nas bases originais)
            # ─────────────────────────────────────────────────────────────────────────────
            def buscar_volume_cliq(row):
                cod = row['Contrato CliqCCEE']
                if cod in ['Verificar', '-', '']: return 0.0
                
                # Procura o código nas bases carregadas
                for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
                    df_cliq = st.session_state.get(db_key)
                    if df_cliq is not None and cod in df_cliq.index:
                        # Pega o valor da coluna MWmedio
                        val = df_cliq.loc[cod, 'MWmedio']
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        return pd.to_numeric(val, errors='coerce')
                return 0.0

            df_conferencia['Volume CliqCCEE'] = df_conferencia.apply(buscar_volume_cliq, axis=1).fillna(0.0).round(6)
            # ─────────────────────────────────────────────────────────────────────────────

            # LÓGICA DE SITUAÇÃO PAGAMENTO
            df_pagos = df_conferencia[df_conferencia['Situacao ERP'].astype(str).str.upper() == 'PAGO'].copy()
            dict_soma_pagos = df_pagos.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()

            def validar_pagamento(row):
                if row['Contrato CliqCCEE'] in ['Verificar', '-', '']: return "-"
                total_pago = dict_soma_pagos.get(row['Contrato CliqCCEE'], 0.0)
                if round(total_pago, 6) == round(row['Volume BOOK'], 6) and row['Volume BOOK'] > 0: return "Pago"
                return "-"
            df_conferencia['SITUAÇÃO PGTO'] = df_conferencia.apply(validar_pagamento, axis=1)

            # FILTROS
            st.write("### Filtros")
            f1, f2, f3, f4, f5, f6 = st.columns([2, 2, 2, 1.2, 1.2, 1.2])
            op_f = f1.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operacao'].unique()))
            parte_f = f2.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique()))
            cliq_f = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].unique()))
            zerar_intra = f4.toggle("Zerar Intraportfólio", value=False)
            zerar_empresas = f5.toggle("Zerar Entre Empresas", value=False)
            ocultar_vazio = f6.toggle("Ocultar Volumes Zerados", value=False)

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if cliq_f != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]
            
            if zerar_intra:
                mask_i = df_final['Vendedor'].str.lower().str.strip() == df_final['Comprador'].str.lower().str.strip()
                df_final.loc[mask_i, ['Montante MWh', 'Volume MWm']] = 0.0

            if zerar_empresas:
                mask_p = (df_final['Parte'].str.contains("BISMUT COMERCIALIZADORA DE ENERGIA S/A", na=False, case=False) |
                          df_final['Parte'].str.contains("GET COMERCIALIZADORA DE ENERGIA S.A.", na=False, case=False))
                mask_c = (df_final['Contraparte'].str.upper().str.startswith("MATRIX", na=False) & 
                          ~df_final['Contraparte'].str.upper().str.contains("MATRIX VAR", na=False))
                df_final.loc[mask_p & mask_c, ['Montante MWh', 'Volume MWm']] = 0.0

            if ocultar_vazio: df_final = df_final[df_final['Volume MWm'] != 0]

            # EXIBIÇÃO
            ordem = [col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP', 
                    'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 
                    'CliqCCEE Paradigma', 'Modulacao WBC', '% Modulacao Min', '% Modulacao Max', 
                    'Contrato CliqCCEE mes anterior', 'Vendedor', 'Comprador', 
                    'Contrato CliqCCEE', 'SITUAÇÃO PGTO', 
                    'Volume BOOK', 'Volume CliqCCEE', # <--- POSIÇÃO SOLICITADA
                    'Situacao ERP', 'Razao Social']
            
            st.dataframe(df_final[ordem].sort_values(by=col_boleta), use_container_width=True, hide_index=True,
                         column_config={
                             "Montante MWh": st.column_config.NumberColumn(format="%.3f"), 
                             "Volume MWm": st.column_config.NumberColumn(format="%.6f"),
                             "Volume BOOK": st.column_config.NumberColumn(format="%.6f"),
                             "Volume CliqCCEE": st.column_config.NumberColumn(format="%.6f") # Formatação para 6 casas
                         })
        else: st.warning("Sem dados para este período.")
    except Exception as e: st.error(f"Erro: {e}")
