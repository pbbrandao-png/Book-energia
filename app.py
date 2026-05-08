import streamlit as st
import pandas as pd
import re
import os
import pickle
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# ─────────────────────────────────────────────────────────────────────────────
# PASTA DE PERSISTÊNCIA
# ─────────────────────────────────────────────────────────────────────────────
PERSIST_DIR = "dados_persistidos"
os.makedirs(PERSIST_DIR, exist_ok=True)

ARQUIVOS_DISCO = {
    'df_bruto':           os.path.join(PERSIST_DIR, 'df_bruto.pkl'),
    'dict_mes_anterior':  os.path.join(PERSIST_DIR, 'dict_mes_anterior.pkl'),
    'dict_comprador':     os.path.join(PERSIST_DIR, 'dict_comprador.pkl'),
    'dict_vendedor':      os.path.join(PERSIST_DIR, 'dict_vendedor.pkl'),
    'dict_mapa':          os.path.join(PERSIST_DIR, 'dict_mapa.pkl'),
    'dict_pendencias':    os.path.join(PERSIST_DIR, 'dict_pendencias.pkl'),
    'db_matrix':          os.path.join(PERSIST_DIR, 'db_matrix.pkl'),
    'db_bismut':          os.path.join(PERSIST_DIR, 'db_bismut.pkl'),
    'db_ccear':           os.path.join(PERSIST_DIR, 'db_ccear.pkl'),
    'db_cbr':             os.path.join(PERSIST_DIR, 'db_cbr.pkl'),
    'ajustes_manuais':    os.path.join(PERSIST_DIR, 'ajustes_manuais.pkl'),
}

def salvar_disco(chave, valor):
    try:
        with open(ARQUIVOS_DISCO[chave], 'wb') as f:
            pickle.dump(valor, f)
    except Exception as e:
        st.warning(f"Não foi possível salvar '{chave}' em disco: {e}")

def carregar_disco(chave, default=None):
    path = ARQUIVOS_DISCO.get(chave)
    if path and os.path.exists(path):
        try:
            with open(path, 'rb') as f:
                return pickle.load(f)
        except Exception:
            return default
    return default

# ─────────────────────────────────────────────────────────────────────────────
# CONTRATOS ESPECIAIS
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
        if df_cliq is None: continue
        if cod in df_cliq.index:
            try:
                mod = df_cliq.loc[cod, 'TIPO_MODULACAO']
                if isinstance(mod, pd.Series): mod = mod.iloc[0]
                if not pd.isna(mod) and str(mod).strip() != "":
                    return str(mod).strip().capitalize()
            except: continue
    return "-"

def buscar_limite_cliq(cod, coluna):
    if cod in ['Verificar', '-', ''] or not cod: return "-"
    if cod in CONTRATOS_ESPECIAIS_CCEAR: return "-"
    for db_key in ['db_matrix', 'db_bismut', 'db_cbr']:
        df_cliq = st.session_state.get(db_key)
        if df_cliq is not None and cod in df_cliq.index and coluna in df_cliq.columns:
            try:
                val = df_cliq.loc[cod, coluna]
                if isinstance(val, pd.Series): val = val.iloc[0]
                if pd.isna(val) or str(val).strip() == "": continue
                return round(float(str(val).replace(',', '.')), 6)
            except: continue
    return "-"

def verificar_match_ccee_linha(vendedor, comprador, submercado_wbc, is_bismut):
    if not vendedor or not comprador or not submercado_wbc: return None, []
    sub_upper  = submercado_wbc.strip().upper()
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
    sem_match  = df_res[df_res['_match'] == False].drop(columns=['_match'])
    com_match  = df_res[df_res['_match'] == True].drop(columns=['_match'])
    incompleto = df_res[df_res['_match'].isna()].drop(columns=['_match'])
    return com_match, sem_match, incompleto

# 4. INICIALIZAÇÃO
if 'dados_carregados_do_disco' not in st.session_state:
    for chave in ['df_bruto', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        if chave not in st.session_state:
            st.session_state[chave] = carregar_disco(chave, default=None)
    for chave in ['dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias']:
        if chave not in st.session_state:
            st.session_state[chave] = carregar_disco(chave, default={})
    st.session_state['ajustes_manuais'] = carregar_disco('ajustes_manuais', default={})
    st.session_state['dados_carregados_do_disco'] = True

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state:
        st.session_state[chave] = None

# 5. SIDEBAR
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

st.sidebar.title("Configurações")
_idx_mes_default = datetime.now().month - 1
_idx_ano_default = anos.index(str(datetime.now().year)) if str(datetime.now().year) in anos else 0

mes_sel = st.sidebar.selectbox("Mês", meses_nomes, index=st.session_state.get('_idx_mes', _idx_mes_default))
ano_sel = st.sidebar.selectbox("Ano", anos,         index=st.session_state.get('_idx_ano', _idx_ano_default))
st.session_state['_idx_mes'] = meses_nomes.index(mes_sel)
st.session_state['_idx_ano'] = anos.index(ano_sel)

st.sidebar.markdown("---")

def status_icon(chave):
    val = st.session_state.get(chave)
    if val is None: return "⬜"
    if isinstance(val, dict) and len(val) == 0: return "⬜"
    return "✅"

st.sidebar.markdown(f"{status_icon('df_bruto')} **1. Contratos Aprovados**")
arquivo_subido = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx', 'xlsm'], key="up_contratos")

st.sidebar.markdown(f"{status_icon('dict_mes_anterior')} **2. Base Mês Anterior**")
arquivo_anterior = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx'], key="up_anterior")

st.sidebar.markdown(f"{status_icon('dict_comprador')} **3. Exportador (4)**")
arquivo_pessoas = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx'], key="up_pessoas")

st.sidebar.markdown(f"{status_icon('dict_mapa')} **4. Mapa Financeiro**")
arquivo_mapa = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx'], key="up_mapa")

st.sidebar.markdown(f"{status_icon('dict_pendencias')} **5. Pendências Financeiras**")
arquivo_pendencias = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx'], key="up_pendencias")

st.sidebar.subheader("Bases Cliq CCEE")
st.sidebar.markdown(f"{status_icon('db_ccear')} **Cliq CCEAR_Q**")
arq_ccear = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx', 'csv'], key="up_ccear")
st.sidebar.markdown(f"{status_icon('db_cbr')} **Cliq CBR Mercado**")
arq_cbr = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx', 'csv'], key="up_cbr")
st.sidebar.markdown(f"{status_icon('db_matrix')} **Cliq Matrix**")
arq_cceal1 = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx', 'csv'], key="up_matrix")
st.sidebar.markdown(f"{status_icon('db_bismut')} **Cliq Bismut**")
arq_cceal2 = st.sidebar.file_uploader("Substituir arquivo", type=['xlsx', 'csv'], key="up_bismut")

if st.sidebar.button("🗑️ Limpar todos os arquivos"):
    import shutil
    shutil.rmtree(PERSIST_DIR, ignore_errors=True)
    os.makedirs(PERSIST_DIR, exist_ok=True)
    for k in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor', 'dict_mapa', 'dict_pendencias', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        st.session_state[k] = {} if 'dict' in k else None
    st.session_state['ajustes_manuais'] = {}
    st.rerun()

st.title(f"Livro de Energia - {mes_sel}/{ano_sel}")

# 6. CARREGAMENTO
if get_file_id(arquivo_subido) != st.session_state.get('fid_subido'):
    st.session_state['fid_subido'] = get_file_id(arquivo_subido)
    if arquivo_subido:
        val = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        st.session_state['df_bruto'] = val
        salvar_disco('df_bruto', val)

if get_file_id(arquivo_anterior) != st.session_state.get('fid_anterior'):
    st.session_state['fid_anterior'] = get_file_id(arquivo_anterior)
    if arquivo_anterior:
        try:
            df_apoio = pd.read_excel(arquivo_anterior, dtype=str)
            val = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].apply(tratar_chave).values).to_dict()
            st.session_state['dict_mes_anterior'] = val
            salvar_disco('dict_mes_anterior', val)
        except: st.session_state['dict_mes_anterior'] = {}

if get_file_id(arquivo_pendencias) != st.session_state.get('fid_pendencias'):
    st.session_state['fid_pendencias'] = get_file_id(arquivo_pendencias)
    if arquivo_pendencias:
        try:
            df_p = pd.read_excel(arquivo_pendencias)
            df_p_simples = df_p.iloc[:, [4, 8]].copy()
            df_p_simples.columns = ['razao_social_pend', 'valor_pendente']
            df_p_simples['valor_pendente'] = pd.to_numeric(df_p_simples['valor_pendente'], errors='coerce').fillna(0)
            df_p_simples['razao_social_pend'] = df_p_simples['razao_social_pend'].astype(str).str.strip().str.upper()
            df_somado = df_p_simples.groupby('razao_social_pend')['valor_pendente'].sum().reset_index()
            val = dict(zip(df_somado['razao_social_pend'], df_somado['valor_pendente']))
            st.session_state['dict_pendencias'] = val
            salvar_disco('dict_pendencias', val)
        except: st.session_state['dict_pendencias'] = {}

if get_file_id(arquivo_pessoas) != st.session_state.get('fid_pessoas'):
    st.session_state['fid_pessoas'] = get_file_id(arquivo_pessoas)
    if arquivo_pessoas:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        val_comp = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        val_vend = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
        st.session_state['dict_comprador'] = val_comp
        st.session_state['dict_vendedor']  = val_vend
        salvar_disco('dict_comprador', val_comp)
        salvar_disco('dict_vendedor',  val_vend)

if get_file_id(arquivo_mapa) != st.session_state.get('fid_mapa'):
    st.session_state['fid_mapa'] = get_file_id(arquivo_mapa)
    if arquivo_mapa:
        df_m = pd.read_excel(arquivo_mapa)
        val = pd.Series(df_m['Situacao_ERP'].values, index=df_m['Codigo_WBC'].apply(tratar_chave).values).to_dict()
        st.session_state['dict_mapa'] = val
        salvar_disco('dict_mapa', val)

if (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)) != st.session_state.get('chave_matrix'):
    st.session_state['chave_matrix'] = (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1))
    if arq_ccear:
        val = carregar_csv_cliq(arq_ccear)
        st.session_state['db_ccear'] = val
        salvar_disco('db_ccear', val)
    if arq_cbr:
        val = carregar_csv_cliq(arq_cbr)
        st.session_state['db_cbr'] = val
        salvar_disco('db_cbr', val)
    if arq_cceal1:
        val = carregar_csv_cliq(arq_cceal1)
        st.session_state['db_matrix'] = val
        salvar_disco('db_matrix', val)

if get_file_id(arq_cceal2) != st.session_state.get('fid_cceal2'):
    st.session_state['fid_cceal2'] = get_file_id(arq_cceal2)
    if arq_cceal2:
        val = carregar_csv_cliq(arq_cceal2)
        st.session_state['db_bismut'] = val
        salvar_disco('db_bismut', val)

# 7. PROCESSAMENTO
if st.session_state['df_bruto'] is not None:
    try:
        st.write("### 🛠️ Ajustes de Boleta")
        with st.expander("Expandir painel de ajustes"):
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
                        salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
                        st.rerun()
            with tab_lote:
                arquivo_lote = st.file_uploader("Planilha de Ajustes", type=['xlsx'], key="upload_lote")
                if arquivo_lote:
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
                        salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
                        st.rerun()

            if st.session_state['ajustes_manuais']:
                for bol_id, dados in list(st.session_state['ajustes_manuais'].items()):
                    col_info, col_del = st.columns([6, 1])
                    col_info.info(f"Boleta {bol_id}")
                    if col_del.button("Remover", key=f"del_{bol_id}"):
                        del st.session_state['ajustes_manuais'][bol_id]
                        salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
                        st.rerun()

        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        mes_num_sel = meses_nomes.index(mes_sel) + 1
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            df_conferencia['Operacao']     = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte']        = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            df_conferencia['Submercado']   = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})
            
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)
            
            df_conferencia['Situacao ERP']       = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            for bol, info in st.session_state['ajustes_manuais'].items():
                mask = df_conferencia['Boleta_Key'] == bol
                if mask.any():
                    df_conferencia.loc[mask, 'Editado'] = True
                    if info['Vendedor']:           df_conferencia.loc[mask, 'Vendedor'] = info['Vendedor']
                    if info['Comprador']:          df_conferencia.loc[mask, 'Comprador'] = info['Comprador']
                    if info['CliqCCEE Paradigma']: df_conferencia.loc[mask, 'CliqCCEE Paradigma'] = info['CliqCCEE Paradigma']

            def resolver_cliq(row):
                vend, comp = row['Vendedor'], row['Comprador']
                if 'BISMUT' in str(row['Parte']).upper():
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear', 'db_ccear'), ('cbr', 'db_cbr'), ('matrix', 'db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)
            
            df_soma_cliq = df_conferencia[~df_conferencia['Contrato CliqCCEE'].isin(['Verificar', '-', ''])].copy()
            dict_soma_book = df_soma_cliq.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['Volume BOOK'] = df_conferencia['Contrato CliqCCEE'].map(dict_soma_book).fillna(0.0).round(6)

            df_conferencia['Volume CliqCCEE'] = df_conferencia.apply(lambda r: buscar_volume_cliq(r), axis=1).fillna(0.0).round(6)
            df_conferencia['Validação Volume'] = df_conferencia.apply(lambda r: "OK" if round(r['Volume BOOK'], 6) == round(r['Volume CliqCCEE'], 6) else "VERIFICAR", axis=1)

            st.write("### Filtros")
            f1, f2, f3, f4 = st.columns(4)
            op_f    = f1.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operacao'].unique()))
            parte_f = f2.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique()))
            cliq_f  = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].unique()))
            valid_vol_f = f4.selectbox("Validação Volume", ["Todos"] + ["OK", "VERIFICAR"])

            df_final = df_conferencia.copy()
            if op_f != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if cliq_f != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]
            if valid_vol_f != "Todos": df_final = df_final[df_final['Validação Volume'] == valid_vol_f]

            st.write("### 📦 Resumo de Operações")
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("🛒 Compras", int((df_final['Operacao'].astype(str).str.upper() == 'COMPRA').sum()))
            mc2.metric("📤 Vendas",  int((df_final['Operacao'].astype(str).str.upper() == 'VENDA').sum()))
            mc3.metric("📋 Total",   len(df_final))

            st.dataframe(
                df_final.sort_values(by=col_boleta),
                use_container_width=True, hide_index=True,
                column_config={
                    "Volume MWm": st.column_config.NumberColumn(format="%.6f"),
                    "Volume BOOK": st.column_config.NumberColumn(format="%.6f"),
                    "Volume CliqCCEE": st.column_config.NumberColumn(format="%.6f"),
                }
            )
        else: st.warning("Sem dados para este período.")
    except Exception as e: st.error(f"Erro: {e}")

def buscar_volume_cliq(row):
    cod = row['Contrato CliqCCEE']
    if cod in ['Verificar', '-', '']: return 0.0
    for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        df_cliq = st.session_state.get(db_key)
        if df_cliq is not None and cod in df_cliq.index:
            val = df_cliq.loc[cod, ('MONTANTE_MENSAL_MWh' if cod in CONTRATOS_ESPECIAIS_CCEAR else 'MWmedio')]
            val = val.iloc[0] if isinstance(val, pd.Series) else val
            return float(str(val).replace(',', '.'))
    return 0.0
