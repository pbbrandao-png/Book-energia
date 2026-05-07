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
        if 'CODIGO_CONTRATO' in df.columns:
            df['CODIGO_CONTRATO'] = df['CODIGO_CONTRATO'].apply(tratar_chave)
            df = df.set_index('CODIGO_CONTRATO')
            return df
        return None
    except Exception:
        return None

# 3. REGRAS DE BUSCA CLIQ
COLUNAS_CLIQ = {
    'matrix': {'vendedor': 'SIGLA_PERFIL_VENDEDOR', 'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'bismut': {'vendedor': 'SIGLA_PERFIL_VENDEDOR', 'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'cbr':    {'vendedor': 'SIGLA_PERFIL_VENDEDOR', 'comprador': 'SIGLA_PERFIL_COMPRADOR'},
    'ccear':  {'vendedor': 'SIGLA_PERFIL_VENDEDOR', 'comprador': 'SIGLA_PERFIL_COMPRADOR'},
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

def buscar_modulacao_cliq(row):
    cod = row['Contrato CliqCCEE']
    if cod in ['Verificar', '-', '']: return "-"
    if cod in CONTRATOS_ESPECIAIS_CCEAR: return "Carga"
    for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        df_cliq = st.session_state.get(db_key)
        if df_cliq is not None and cod in df_cliq.index:
            try:
                mod = df_cliq.loc[cod, 'TIPO_MODULACAO']
                if isinstance(mod, pd.Series): mod = mod.iloc[0]
                if not pd.isna(mod) and str(mod).strip() != "":
                    return str(mod).strip().capitalize()
            except: continue
    return "-"

def verificar_match_ccee_linha(vendedor, comprador, submercado_wbc, is_bismut):
    if not vendedor or not comprador or not submercado_wbc: return None, []
    sub_upper = submercado_wbc.strip().upper()
    vend_upper = vendedor.strip().upper()
    comp_upper = comprador.strip().upper()
    bases = ['db_bismut'] if is_bismut else ['db_ccear', 'db_cbr', 'db_matrix']
    bases_consultadas = []
    for db_key in bases:
        df_cliq = st.session_state.get(db_key)
        if df_cliq is None: continue
        df_temp = df_cliq.reset_index()
        bases_consultadas.append(db_key.replace('db_', '').upper())
        mask = (
            (df_temp['SUBMERCADO_ENTREGA'].astype(str).str.strip().str.upper() == sub_upper) &
            (df_temp['SIGLA_PERFIL_VENDEDOR'].astype(str).str.strip().str.upper() == vend_upper) &
            (df_temp['SIGLA_PERFIL_COMPRADOR'].astype(str).str.strip().str.upper() == comp_upper)
        )
        if mask.any(): return True, bases_consultadas
    return False, bases_consultadas

def gerar_relatorio_match(df_conferencia):
    resultados = []
    for _, row in df_conferencia.iterrows():
        volume = pd.to_numeric(row.get('Volume MWm', 0), errors='coerce')
        if pd.isna(volume) or volume == 0: continue
        vendedor   = str(row.get('Vendedor', '')).strip() if row.get('Vendedor', '-') != '-' else ''
        comprador  = str(row.get('Comprador', '')).strip() if row.get('Comprador', '-') != '-' else ''
        submercado = str(row.get('Submercado', '')).strip()
        is_bismut  = 'BISMUT' in str(row.get('Parte', '')).upper()
        boleta     = row.iloc[0] if len(row) > 0 else ''
        match, bases = verificar_match_ccee_linha(vendedor, comprador, submercado, is_bismut)
        resultados.append({
            'Boleta': boleta, 'Parte': row.get('Parte', ''), 'Contraparte': row.get('Contraparte', ''),
            'Submercado': submercado, 'Vendedor': vendedor, 'Comprador': comprador,
            'Bases Consultadas': ', '.join(bases) if bases else '-', '_match': match
        })
    df_res = pd.DataFrame(resultados)
    if df_res.empty: return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
    sem_match = df_res[df_res['_match'] == False].drop(columns=['_match'])
    com_match = df_res[df_res['_match'] == True].drop(columns=['_match'])
    incompleto = df_res[df_res['_match'].isna()].drop(columns=['_match'])
    return com_match, sem_match, incompleto

# 4. INICIALIZAÇÃO
if 'ajustes_manuais' not in st.session_state:
    st.session_state['ajustes_manuais'] = {}

meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

st.sidebar.title("Configurações")
mes_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
ano_sel = st.sidebar.selectbox("Ano", anos, index=anos.index(str(datetime.now().year)))

st.title(f"Livro de Energia - {mes_sel}/{ano_sel}")

st.session_state['mes_sel'] = mes_sel
st.session_state['ano_sel'] = ano_sel

st.sidebar.markdown("---")
arquivo_subido     = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior   = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx", type=['xlsx'])
arquivo_pessoas    = st.sidebar.file_uploader("3. Exportador (4).xlsx", type=['xlsx'])
arquivo_mapa       = st.sidebar.file_uploader("4. Mapa Financeiro (Excel)", type=['xlsx'])
arquivo_pendencias = st.sidebar.file_uploader("5. Pendências Financeiras (Excel)", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q", type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq Matrix", type=['xlsx', 'csv'])
arq_cceal2 = st.sidebar.file_uploader("Cliq Bismut", type=['xlsx', 'csv'])

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias',
              'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
    if chave not in st.session_state: st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state: st.session_state[chave] = None

# 6. CARREGAMENTO DOS DADOS
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
            df_p_simples = df_p.iloc[:, [4, 8]].copy(); df_p_simples.columns = ['razao_social_pend', 'valor_pendente']
            df_p_simples['valor_pendente'] = pd.to_numeric(df_p_simples['valor_pendente'], errors='coerce').fillna(0)
            df_p_simples['razao_social_pend'] = df_p_simples['razao_social_pend'].astype(str).str.strip().str.upper()
            df_somado = df_p_simples.groupby('razao_social_pend')['valor_pendente'].sum().reset_index()
            st.session_state['dict_pendencias'] = dict(zip(df_somado['razao_social_pend'], df_somado['valor_pendente']))
        except: st.session_state['dict_pendencias'] = {}

if get_file_id(arquivo_pessoas) != st.session_state['fid_pessoas']:
    st.session_state['fid_pessoas'] = get_file_id(arquivo_pessoas)
    if arquivo_pessoas:
        df_pers = pd.read_excel(arquivo_pessoas); df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        st.session_state['dict_comprador'] = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        st.session_state['dict_vendedor']  = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()

if get_file_id(arquivo_mapa) != st.session_state['fid_mapa']:
    st.session_state['fid_mapa'] = get_file_id(arquivo_mapa)
    if arquivo_mapa:
        df_m = pd.read_excel(arquivo_mapa)
        st.session_state['dict_mapa'] = pd.Series(df_m['Situacao_ERP'].values, index=df_m['Codigo_WBC'].apply(tratar_chave).values).to_dict()

if (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)) != st.session_state['chave_matrix']:
    st.session_state['chave_matrix'] = (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1))
    st.session_state['db_ccear'] = carregar_csv_cliq(arq_ccear)
    st.session_state['db_cbr'] = carregar_csv_cliq(arq_cbr); st.session_state['db_matrix'] = carregar_csv_cliq(arq_cceal1)

if get_file_id(arq_cceal2) != st.session_state['fid_cceal2']:
    st.session_state['fid_cceal2'] = get_file_id(arq_cceal2)
    st.session_state['db_bismut'] = carregar_csv_cliq(arq_cceal2)

# 7. PROCESSAMENTO DA TABELA
if st.session_state['df_bruto'] is not None:
    try:
        # --- SEÇÃO DE AJUSTES MANUAIS ---
        st.write("### 🛠️ Ajustes de Boleta")
        with st.expander("Expandir painel de ajustes (Individual ou Lote)"):
            tab_manual, tab_lote = st.tabs(["Edição Individual", "Upload em Lote"])
            
            with tab_manual:
                c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
                edit_bol  = c1.text_input("ID Boleta", key="input_bol")
                edit_vend = c2.text_input("Novo Vendedor", key="input_vend")
                edit_comp = c3.text_input("Novo Comprador", key="input_comp")
                edit_cliq = c4.text_input("Novo Cliq Paradigma", key="input_cliq")
                
                if st.button("Gravar Alteração"):
                    if edit_bol:
                        st.session_state['ajustes_manuais'][tratar_chave(edit_bol)] = {
                            'Vendedor': edit_vend if edit_vend else None,
                            'Comprador': edit_comp if edit_comp else None,
                            'CliqCCEE Paradigma': edit_cliq if edit_cliq else None
                        }
                        st.success(f"Boleta {edit_bol} atualizada!")
                        st.rerun()
            
            with tab_lote:
                st.write("Suba uma planilha com as colunas: **BOLETA**, **Vendedor**, **Comprador**, **CliqCCEE Paradigma**")
                arquivo_lote = st.file_uploader("Planilha de Ajustes", type=['xlsx'], key="upload_lote")
                if arquivo_lote:
                    try:
                        df_lote = pd.read_excel(arquivo_lote)
                        df_lote.columns = [c.strip() for c in df_lote.columns]
                        if 'BOLETA' in df_lote.columns:
                            for _, r in df_lote.iterrows():
                                b_id = tratar_chave(r['BOLETA'])
                                if b_id:
                                    st.session_state['ajustes_manuais'][b_id] = {
                                        'Vendedor': str(r['Vendedor']).strip() if 'Vendedor' in r and not pd.isna(r['Vendedor']) else None,
                                        'Comprador': str(r['Comprador']).strip() if 'Comprador' in r and not pd.isna(r['Comprador']) else None,
                                        'CliqCCEE Paradigma': tratar_chave(r['CliqCCEE Paradigma']) if 'CliqCCEE Paradigma' in r and not pd.isna(r['CliqCCEE Paradigma']) else None
                                    }
                            st.success("Ajustes em lote carregados com sucesso!")
                            st.rerun()
                        else:
                            st.error("Coluna 'BOLETA' não encontrada no arquivo.")
                    except Exception as e_lote:
                        st.error(f"Erro ao processar lote: {e_lote}")

            if st.session_state['ajustes_manuais']:
                st.markdown("---")
                st.write("**Ajustes Ativos:**")
                for bol_id, dados in list(st.session_state['ajustes_manuais'].items()):
                    col_info, col_del = st.columns([6, 1])
                    info_parts = [f"ID: {bol_id}"]
                    if dados['Vendedor']: info_parts.append(f"Vend: {dados['Vendedor']}")
                    if dados['Comprador']: info_parts.append(f"Comp: {dados['Comprador']}")
                    if dados['CliqCCEE Paradigma']: info_parts.append(f"Cliq: {dados['CliqCCEE Paradigma']}")
                    
                    col_info.info(" | ".join(info_parts))
                    if col_del.button("Remover", key=f"del_{bol_id}"):
                        del st.session_state['ajustes_manuais'][bol_id]
                        st.rerun()

                if st.button("Limpar todos os ajustes"):
                    st.session_state['ajustes_manuais'] = {}
                    st.rerun()

        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        
        mes_num_sel = meses_nomes.index(st.session_state['mes_sel']) + 1
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()

            mapa_energia = {'Incentivada-50%': 'Incentivada-I5', 'Incentivada-100%': 'Incentivada-I1', 'Incentivada-0%': 'Incentivada-I0', 'Incentivada-CQ50%': 'Incentivada-CQ5'}
            df_conferencia['Tipo Energia'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]]).astype(str).str.strip().replace(mapa_energia)

            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})

            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes_serie = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes_serie).fillna(0).round(6)

            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            df_conferencia['% Modulacao Min'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[28]]), errors='coerce').fillna("-")
            df_conferencia['% Modulacao Max'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[29]]), errors='coerce').fillna("-")
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            # --- APLICAR AJUSTES MANUAIS ---
            df_conferencia['Editado'] = False
            for bol, info in st.session_state['ajustes_manuais'].items():
                mask = df_conferencia['Boleta_Key'] == bol
                if mask.any():
                    df_conferencia.loc[mask, 'Editado'] = True
                    if info['Vendedor']: df_conferencia.loc[mask, 'Vendedor'] = info['Vendedor']
                    if info['Comprador']: df_conferencia.loc[mask, 'Comprador'] = info['Comprador']
                    if info['CliqCCEE Paradigma']: df_conferencia.loc[mask, 'CliqCCEE Paradigma'] = info['CliqCCEE Paradigma']

            def resolver_cliq(row):
                vend, comp = (row['Vendedor'] if row['Vendedor'] != "-" else ""), (row['Comprador'] if row['Comprador'] != "-" else "")
                if 'BISMUT' in str(row['Parte']).upper(): return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear', 'db_ccear'), ('cbr', 'db_cbr'), ('matrix', 'db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)
            df_conferencia['Modulação CCEE'] = df_conferencia.apply(buscar_modulacao_cliq, axis=1)

            def buscar_status_cliq(row):
                cod = row['Contrato CliqCCEE']
                if cod in ['Verificar', '-', '']: return "-"
                for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
                    df_cliq = st.session_state.get(db_key)
                    if df_cliq is not None and cod in df_cliq.index:
                        status = df_cliq.loc[cod, 'SITUACAO_CONTRATO']
                        return str(status.iloc[0] if isinstance(status, pd.Series) else status).strip()
                return "-"

            df_conferencia['Status do Contrato'] = df_conferencia.apply(buscar_status_cliq, axis=1)
            df_soma_cliq = df_conferencia[~df_conferencia['Contrato CliqCCEE'].isin(['Verificar', '-', ''])].copy()
            dict_soma_book = df_soma_cliq.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['Volume BOOK'] = df_conferencia['Contrato CliqCCEE'].map(dict_soma_book).fillna(0.0).round(6)

            # MELHORIA: VOLUME CLIQ CCEE COM PROTEÇÃO DE HORAS
            def buscar_volume_cliq(row):
                cod = row['Contrato CliqCCEE']
                if cod in ['Verificar', '-', '']: return 0.0
                
                # Pega as horas do mês da boleta atual
                h_mes_row = h_mes_serie.get(row.name, 744)
                if pd.isna(h_mes_row) or h_mes_row == 0: h_mes_row = 744

                for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
                    df_cliq = st.session_state.get(db_key)
                    if df_cliq is not None and cod in df_cliq.index:
                        try:
                            # Se for CCEAR_Q usa MONTANTE_MENSAL_MWh
                            if cod in CONTRATOS_ESPECIAIS_CCEAR:
                                val = df_cliq.loc[cod, 'MONTANTE_MENSAL_MWh']
                                val = val.iloc[0] if isinstance(val, pd.Series) else val
                                v = float(str(val).replace(',', '.'))
                                return v / h_mes_row
                            else:
                                # Senão usa MWmedio padrão
                                val = df_cliq.loc[cod, 'MWmedio']
                                val = val.iloc[0] if isinstance(val, pd.Series) else val
                                return float(str(val).replace(',', '.'))
                        except: continue
                return 0.0

            df_conferencia['Volume CliqCCEE'] = df_conferencia.apply(buscar_volume_cliq, axis=1).fillna(0.0).round(6)
            
            def validar_volume_logic(row):
                if row['Contrato CliqCCEE'] in ['Verificar', '-', '']: return "-"
                return "OK" if round(row['Volume BOOK'], 6) == round(row['Volume CliqCCEE'], 6) else "VERIFICAR"

            df_conferencia['Validação Volume'] = df_conferencia.apply(validar_volume_logic, axis=1)

            # --- PROTEÇÃO PARA COLUNAS FINANCEIRAS ---
            df_pagos = df_conferencia[df_conferencia['Situacao ERP'].astype(str).str.upper() == 'PAGO'].copy()
            dict_soma_pagos = df_pagos.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()

            def validar_pagamento(row):
                if row['Contrato CliqCCEE'] in ['Verificar', '-', '']: return "-"
                total_pago = dict_soma_pagos.get(row['Contrato CliqCCEE'], 0.0)
                return "Pago" if round(total_pago, 6) >= round(row['Volume BOOK'], 6) and row['Volume BOOK'] > 0 else "-"

            # Criando colunas de forma segura para não quebrar o index
            df_conferencia['SITUAÇÃO PGTO'] = df_conferencia.apply(validar_pagamento, axis=1)
            
            if st.session_state['dict_pendencias']:
                df_conferencia['Pendência Financeira'] = df_conferencia['Razao Social'].str.strip().str.upper().map(st.session_state['dict_pendencias']).fillna(0.0)
            else:
                df_conferencia['Pendência Financeira'] = 0.0

            st.write("### Filtros")
            f1, f2, f3, f4, f5, f6, f7 = st.columns([1.5, 1.5, 1.5, 1.5, 1.2, 1.2, 1.2])
            op_f = f1.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operacao'].unique()))
            parte_f = f2.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique()))
            cliq_f = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].unique()))
            valid_vol_f = f4.selectbox("Validação Volume", ["Todos"] + sorted(df_conferencia['Validação Volume'].unique()))
            zerar_intra, zerar_entre, ocultar_vazio = f5.toggle("Zerar Intraportfólio"), f6.toggle("Zerar Entre Empresas"), f7.toggle("Ocultar Volumes Zerados")

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if cliq_f != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]
            if valid_vol_f != "Todos": df_final = df_final[df_final['Validação Volume'] == valid_vol_f]
            
            if zerar_intra:
                mask_i = df_final['Vendedor'].str.lower().str.strip() == df_final['Comprador'].str.lower().str.strip()
                df_final.loc[mask_i, ['Montante MWh', 'Volume MWm']] = 0.0
            if zerar_entre:
                mask_p = df_final['Parte'].str.contains("BISMUT|GET", na=False, case=False)
                mask_c = df_final['Contraparte'].str.upper().str.startswith("MATRIX", na=False) & ~df_final['Contraparte'].str.upper().str.contains("MATRIX VAR", na=False)
                df_final.loc[mask_p & mask_c, ['Montante MWh', 'Volume MWm']] = 0.0
            if ocultar_vazio: df_final = df_final[df_final['Volume MWm'] != 0]

            # ─────────────────────────────────────────────────────────────────
            # BALÕES DE MÉTRICAS (METRICS)
            # ─────────────────────────────────────────────────────────────────
            st.write("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Contratos Filtrados", len(df_final))
            m2.metric("Compras", len(df_final[df_final['Operacao'].str.upper() == 'COMPRA']))
            m3.metric("Vendas", len(df_final[df_final['Operacao'].str.upper() == 'VENDA']))
            
            # Conta quantos contratos distintos do Cliq CCEE foram identificados no filtro
            cliqs_unicos = df_final[~df_final['Contrato CliqCCEE'].isin(['Verificar', '-', ''])]['Contrato CliqCCEE'].nunique()
            m4.metric("Contratos Cliq identificados", cliqs_unicos)

            bases_carregadas = [k.replace('db_', '').upper() for k in ['db_ccear', 'db_cbr', 'db_matrix', 'db_bismut'] if st.session_state.get(k) is not None]
            with st.expander(f"🔍 Validação de Match CCEE", expanded=False):
                if not bases_carregadas: st.warning("Nenhuma base Cliq CCEE carregada.")
                else:
                    _, sem_match, _ = gerar_relatorio_match(df_final)
                    if sem_match.empty: st.success("Nenhuma linha sem match!")
                    else: 
                        st.warning(f"{len(sem_match)} linha(s) sem contrato correspondente.")
                        st.dataframe(sem_match.reset_index(drop=True), use_container_width=True, hide_index=True)

            # --- ESTILIZAÇÃO E EXIBIÇÃO ---
            ordem = [col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP', 'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 'CliqCCEE Paradigma', 'Modulacao WBC', 'Modulação CCEE', '% Modulacao Min', '% Modulacao Max', 'Contrato CliqCCEE mes anterior', 'Vendedor', 'Comprador', 'Contrato CliqCCEE', 'Status do Contrato', 'SITUAÇÃO PGTO', 'Volume BOOK', 'Volume CliqCCEE', 'Validação Volume', 'Situacao ERP', 'Razao Social', 'Pendência Financeira', 'Editado']
            
            # Garante que só tentará exibir colunas que realmente existem no DataFrame
            ordem_final = [c for c in ordem if c in df_final.columns]

            def highlight_rows(row):
                if row.get('Editado', False):
                    return ['background-color: #fff4cc'] * len(row)
                return [''] * len(row)

            st.dataframe(
                df_final[ordem_final].sort_values(by=col_boleta).style.apply(highlight_rows, axis=1), 
                use_container_width=True, 
                hide_index=True, 
                column_config={
                    "Editado": None,
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"), 
                    "Volume MWm": st.column_config.NumberColumn(format="%.6f"), 
                    "Volume BOOK": st.column_config.NumberColumn(format="%.6f"), 
                    "Volume CliqCCEE": st.column_config.NumberColumn(format="%.6f"), 
                    "Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f")
                }
            )
        else: st.warning("Sem dados para este período.")
    except Exception as e: st.error(f"Erro no processamento: {e}")
