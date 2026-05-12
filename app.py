import streamlit as st
import pandas as pd
import re
import os
import pickle
import zipfile
import io
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# ─────────────────────────────────────────────────────────────────────────────
# PASTA DE PERSISTÊNCIA — cria se não existir
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

# ─────────────────────────────────────────────────────────────────────────────
# ZIP BROWSER INLINE
# Renderiza navegador de ZIP diretamente num container (sidebar ou main).
# Retorna um BytesIO com .name e .size quando o usuário selecionar um arquivo,
# ou None enquanto nenhum arquivo estiver selecionado.
# ─────────────────────────────────────────────────────────────────────────────
EXTS_PLANILHA = ('.xlsx', '.xlsm', '.csv')

def _listar_zip(zf, pasta_atual):
    subpastas = set()
    arquivos  = []
    for nome in zf.namelist():
        if not nome.startswith(pasta_atual):
            continue
        resto = nome[len(pasta_atual):]
        if not resto:
            continue
        partes = resto.split('/')
        if len(partes) == 1:
            if partes[0].lower().endswith(EXTS_PLANILHA):
                arquivos.append(nome)
        else:
            subpastas.add(pasta_atual + partes[0] + '/')
    return sorted(subpastas), sorted(arquivos)


def zip_browser_inline(zip_file_obj, state_prefix, container):
    """
    Renderiza o navegador de ZIP dentro de `container` (st.sidebar ou st).
    - zip_file_obj : objeto retornado pelo file_uploader (já lido)
    - state_prefix : prefixo único no session_state para este uploader
    - container    : onde renderizar (st.sidebar ou st)

    Retorna BytesIO do arquivo selecionado, ou None.
    """
    zip_id = (zip_file_obj.name, zip_file_obj.size)
    key_id    = f"{state_prefix}_zip_id"
    key_pasta = f"{state_prefix}_zip_pasta"
    key_sel   = f"{state_prefix}_zip_selected"
    key_bytes = f"{state_prefix}_zip_bytes"

    # Reinicia navegação se o zip mudou
    if st.session_state.get(key_id) != zip_id:
        st.session_state[key_id]    = zip_id
        st.session_state[key_pasta] = ''
        st.session_state[key_sel]   = None
        st.session_state[key_bytes] = zip_file_obj.read()

    zf = zipfile.ZipFile(io.BytesIO(st.session_state[key_bytes]))
    pasta_atual = st.session_state.get(key_pasta, '')
    subpastas, arquivos = _listar_zip(zf, pasta_atual)

    # Cabeçalho de navegação
    if pasta_atual:
        partes = pasta_atual.rstrip('/').split('/')
        pasta_pai = '/'.join(partes[:-1])
        if pasta_pai:
            pasta_pai += '/'
        container.caption(f"📂 `/{pasta_atual}`")
        if container.button("⬆️ Voltar", key=f"{state_prefix}_voltar"):
            st.session_state[key_pasta] = pasta_pai
            st.session_state[key_sel]   = None
            st.rerun()
    else:
        container.caption(f"📂 raiz de `{zip_file_obj.name}`")

    # Pastas
    for pasta in subpastas:
        nome_exib = pasta[len(pasta_atual):].strip('/')
        if container.button(f"📁 {nome_exib}", key=f"{state_prefix}_pasta_{pasta}"):
            st.session_state[key_pasta] = pasta
            st.session_state[key_sel]   = None
            st.rerun()

    # Arquivos de planilha
    if arquivos:
        for arq_path in arquivos:
            nome_exib  = arq_path[len(pasta_atual):]
            selecionado = st.session_state.get(key_sel) == arq_path
            label = f"{'✅' if selecionado else '📄'} {nome_exib}"
            if container.button(label, key=f"{state_prefix}_arq_{arq_path}"):
                st.session_state[key_sel] = arq_path
                st.rerun()
    elif not subpastas:
        container.info("Nenhuma planilha encontrada nesta pasta.")

    # Retorna BytesIO se houver seleção
    sel = st.session_state.get(key_sel)
    if sel:
        nome_curto = sel.split('/')[-1]
        container.success(f"✅ Selecionado: **{nome_curto}**")
        data = zf.read(sel)
        buf  = io.BytesIO(data)
        buf.name = nome_curto
        buf.size = len(data)
        return buf

    return None


def uploader_com_zip(label, types, key, state_prefix, container):
    """
    Substitui file_uploader: aceita xlsx/xlsm/csv/zip.
    Se o arquivo for ZIP, exibe navegador inline e retorna o arquivo interno
    selecionado (BytesIO). Caso contrário retorna o arquivo diretamente.
    """
    tipos_aceitos = list(set(types + ['zip']))
    arq = container.file_uploader(label, type=tipos_aceitos, key=key)

    if arq is None:
        return None

    if arq.name.lower().endswith('.zip'):
        with container.expander("📦 Navegar dentro do ZIP", expanded=True):
            return zip_browser_inline(arq, state_prefix, st.sidebar if container is st.sidebar else st)
    else:
        return arq


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

# ─────────────────────────────────────────────────────────────────────────────
# 4. INICIALIZAÇÃO
# ─────────────────────────────────────────────────────────────────────────────
if 'dados_carregados_do_disco' not in st.session_state:
    for chave in ['df_bruto', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        if chave not in st.session_state:
            st.session_state[chave] = carregar_disco(chave, default=None)

    for chave in ['dict_mes_anterior', 'dict_comprador', 'dict_vendedor',
                  'dict_mapa', 'dict_pendencias']:
        if chave not in st.session_state:
            st.session_state[chave] = carregar_disco(chave, default={})

    st.session_state['ajustes_manuais'] = carregar_disco('ajustes_manuais', default={})
    st.session_state['dados_carregados_do_disco'] = True

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix',
              'fid_cceal2', 'fid_mapa', 'fid_pendencias']:
    if chave not in st.session_state:
        st.session_state[chave] = None

# ─────────────────────────────────────────────────────────────────────────────
# 5. SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

st.sidebar.title("Configurações")

if '_idx_mes' not in st.session_state:
    st.session_state['_idx_mes'] = datetime.now().month - 1
if '_idx_ano' not in st.session_state:
    st.session_state['_idx_ano'] = anos.index(str(datetime.now().year)) if str(datetime.now().year) in anos else 0

mes_sel = st.sidebar.selectbox("Mês", meses_nomes, index=st.session_state['_idx_mes'], key="sel_mes")
ano_sel = st.sidebar.selectbox("Ano", anos,         index=st.session_state['_idx_ano'], key="sel_ano")
st.session_state['_idx_mes'] = meses_nomes.index(mes_sel)
st.session_state['_idx_ano'] = anos.index(ano_sel)

st.sidebar.markdown("---")

def status_icon(chave):
    val = st.session_state.get(chave)
    if val is None: return "⬜"
    if isinstance(val, dict) and len(val) == 0: return "⬜"
    return "✅"

# ── Uploaders com suporte nativo a ZIP ───────────────────────────────────────

st.sidebar.markdown(f"{status_icon('df_bruto')} **1. Contratos Aprovados**")
arquivo_subido = uploader_com_zip(
    "Substituir arquivo", ['xlsx', 'xlsm'],
    key="up_contratos", state_prefix="zip_contratos", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('dict_mes_anterior')} **2. Base Mês Anterior**")
arquivo_anterior = uploader_com_zip(
    "Substituir arquivo", ['xlsx'],
    key="up_anterior", state_prefix="zip_anterior", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('dict_comprador')} **3. Exportador (4)**")
arquivo_pessoas = uploader_com_zip(
    "Substituir arquivo", ['xlsx'],
    key="up_pessoas", state_prefix="zip_pessoas", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('dict_mapa')} **4. Mapa Financeiro**")
arquivo_mapa = uploader_com_zip(
    "Substituir arquivo", ['xlsx'],
    key="up_mapa", state_prefix="zip_mapa", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('dict_pendencias')} **5. Pendências Financeiras**")
arquivo_pendencias = uploader_com_zip(
    "Substituir arquivo", ['xlsx'],
    key="up_pendencias", state_prefix="zip_pendencias", container=st.sidebar
)

st.sidebar.subheader("Bases Cliq CCEE")

st.sidebar.markdown(f"{status_icon('db_ccear')} **Cliq CCEAR_Q**")
arq_ccear = uploader_com_zip(
    "Substituir arquivo", ['xlsx', 'csv'],
    key="up_ccear", state_prefix="zip_ccear", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('db_cbr')} **Cliq CBR Mercado**")
arq_cbr = uploader_com_zip(
    "Substituir arquivo", ['xlsx', 'csv'],
    key="up_cbr", state_prefix="zip_cbr", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('db_matrix')} **Cliq Matrix**")
arq_cceal1 = uploader_com_zip(
    "Substituir arquivo", ['xlsx', 'csv'],
    key="up_matrix", state_prefix="zip_matrix", container=st.sidebar
)

st.sidebar.markdown(f"{status_icon('db_bismut')} **Cliq Bismut**")
arq_cceal2 = uploader_com_zip(
    "Substituir arquivo", ['xlsx', 'csv'],
    key="up_bismut", state_prefix="zip_bismut", container=st.sidebar
)

if st.sidebar.button("🗑️ Limpar todos os arquivos salvos"):
    import shutil
    shutil.rmtree(PERSIST_DIR, ignore_errors=True)
    os.makedirs(PERSIST_DIR, exist_ok=True)
    for k in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor',
              'dict_mapa', 'dict_pendencias', 'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
        st.session_state[k] = {} if 'dict' in k else None
    st.session_state['ajustes_manuais'] = {}
    st.rerun()

st.title(f"Livro de Energia - {mes_sel}/{ano_sel}")

# ─────────────────────────────────────────────────────────────────────────────
# 6. CARREGAMENTO E PERSISTÊNCIA DOS DADOS
# ─────────────────────────────────────────────────────────────────────────────
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

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS COMPARTILHADOS PARA AS TABELAS CCEE/WBC
# ─────────────────────────────────────────────────────────────────────────────

def adicionar_total(df):
    total = {
        'PERFIL':    'TOTAL',
        'Comprador': round(df['Comprador'].sum(), 6),
        'Vendedor':  round(df['Vendedor'].sum(), 6),
        'NET':       round(df['NET'].sum(), 6),
    }
    return pd.concat([df, pd.DataFrame([total])], ignore_index=True)

def highlight_total(row):
    if row['PERFIL'] == 'TOTAL':
        return ['font-weight: bold; background-color: #e8f0fe'] * len(row)
    return [''] * len(row)

COL_CONFIG_PERFIL = {
    "PERFIL":    st.column_config.TextColumn("PERFIL"),
    "Comprador": st.column_config.NumberColumn("Comprador", format="%.6f"),
    "Vendedor":  st.column_config.NumberColumn("Vendedor",  format="%.6f"),
    "NET":       st.column_config.NumberColumn("NET",       format="%.6f"),
}

def calcular_mwmedio_em_bases(perfil, coluna_perfil, filtro_sub, db_keys):
    total = 0.0
    for db_key in db_keys:
        df_cliq = st.session_state.get(db_key)
        if df_cliq is None:
            continue
        df_temp = df_cliq.reset_index()
        if coluna_perfil not in df_temp.columns or 'MWmedio' not in df_temp.columns:
            continue
        mask = df_temp[coluna_perfil].astype(str).str.strip().str.upper() == perfil.upper()
        if filtro_sub != "Todos" and 'SUBMERCADO_ENTREGA' in df_temp.columns:
            mask = mask & (df_temp['SUBMERCADO_ENTREGA'].astype(str).str.strip() == filtro_sub)
        df_f = df_temp[mask].copy()
        df_f['MWmedio'] = pd.to_numeric(
            df_f['MWmedio'].astype(str).str.strip().str.replace(',', '.'), errors='coerce'
        ).fillna(0.0)
        total += df_f['MWmedio'].sum()
    return round(total, 6)

def build_ccee_tabela(perfis, db_keys, filtro_sub):
    rows = []
    for perfil in perfis:
        comp = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_COMPRADOR', filtro_sub, db_keys)
        vend = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_VENDEDOR',  filtro_sub, db_keys)
        rows.append({'PERFIL': perfil, 'Comprador': comp, 'Vendedor': vend, 'NET': round(comp - vend, 6)})
    return adicionar_total(pd.DataFrame(rows))

def build_wbc_tabela(perfis, df_wbc_base):
    rows = []
    for perfil in perfis:
        mask_c = df_wbc_base['Comprador'].astype(str).str.strip().str.upper() == perfil.upper()
        mask_v = df_wbc_base['Vendedor'].astype(str).str.strip().str.upper()  == perfil.upper()
        soma_c = round(pd.to_numeric(df_wbc_base.loc[mask_c, 'Volume MWm'], errors='coerce').fillna(0.0).sum(), 6)
        soma_v = round(pd.to_numeric(df_wbc_base.loc[mask_v, 'Volume MWm'], errors='coerce').fillna(0.0).sum(), 6)
        rows.append({'PERFIL': perfil, 'Comprador': soma_c, 'Vendedor': soma_v, 'NET': round(soma_c - soma_v, 6)})
    return adicionar_total(pd.DataFrame(rows))

def render_tabela_par(titulo_ccee, titulo_wbc, df_ccee, df_wbc, aviso_sem_base=None):
    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(f"**{titulo_ccee}**")
        if aviso_sem_base:
            st.warning(aviso_sem_base)
        else:
            st.dataframe(
                df_ccee.style.apply(highlight_total, axis=1),
                use_container_width=True, hide_index=True,
                column_config=COL_CONFIG_PERFIL
            )
    with col_r:
        st.markdown(f"**{titulo_wbc}**")
        st.dataframe(
            df_wbc.style.apply(highlight_total, axis=1),
            use_container_width=True, hide_index=True,
            column_config=COL_CONFIG_PERFIL
        )

def render_reconciliacao(titulo, df_ccee, df_wbc):
    df_c = df_ccee[df_ccee['PERFIL'] != 'TOTAL'].copy()
    df_w = df_wbc[df_wbc['PERFIL'] != 'TOTAL'].copy()

    df_rec = df_c[['PERFIL', 'NET']].rename(columns={'NET': 'NET_CCEE'}).merge(
        df_w[['PERFIL', 'NET']].rename(columns={'NET': 'NET_WBC'}),
        on='PERFIL', how='outer'
    ).fillna(0.0)

    df_rec['DIFERENÇA'] = (df_rec['NET_CCEE'] - df_rec['NET_WBC']).round(6)
    df_rec['STATUS'] = df_rec['DIFERENÇA'].apply(
        lambda d: "✅ OK" if abs(d) < 1e-5 else "⚠️ Verificar"
    )

    causas = []
    for _, row in df_rec[df_rec['STATUS'] != "✅ OK"].iterrows():
        perfil = row['PERFIL']
        diff   = row['DIFERENÇA']
        if abs(diff) > 0:
            causas.append(
                f"**{perfil}**: NET CCEE = {row['NET_CCEE']:.6f} | NET WBC = {row['NET_WBC']:.6f} | "
                f"Diferença = **{diff:.6f}**"
            )

    with st.expander(f"🔎 Reconciliação NET — {titulo}", expanded=False):
        st.dataframe(
            df_rec.style.apply(
                lambda r: ['background-color: #fdecea' if r['STATUS'] != '✅ OK' else '' for _ in r],
                axis=1
            ),
            use_container_width=True, hide_index=True,
            column_config={
                "PERFIL":     st.column_config.TextColumn("PERFIL"),
                "NET_CCEE":   st.column_config.NumberColumn("NET CCEE",  format="%.6f"),
                "NET_WBC":    st.column_config.NumberColumn("NET WBC",   format="%.6f"),
                "DIFERENÇA":  st.column_config.NumberColumn("Diferença", format="%.6f"),
                "STATUS":     st.column_config.TextColumn("Status"),
            }
        )
        if causas:
            st.markdown("**Perfis com divergência:**")
            for c in causas:
                st.markdown(f"- {c}")
            st.markdown(
                "_Possíveis causas: boleta não cadastrada na base Cliq, contrato com status RASCUNHO, "
                "operação intraportfólio não zerada, entre empresas não zerada, ou contrato sem match CCEE._"
            )
        else:
            st.success("Todos os perfis estão reconciliados (NET CCEE = NET WBC).")

# ─────────────────────────────────────────────────────────────────────────────
# FUNÇÃO AUXILIAR: classifica Varejista com base no Comprador
# ─────────────────────────────────────────────────────────────────────────────
def classificar_varejista(comprador):
    if pd.isna(comprador) or str(comprador).strip() in ['-', '']:
        return "Não"
    comp_upper = str(comprador).strip().upper()
    if comp_upper.startswith("MATRIX VAR") or comp_upper.startswith("BISMUT VAR"):
        return "Sim"
    return "Não"

def tratar_modulacao_pct(valor_raw):
    v = pd.to_numeric(valor_raw, errors='coerce')
    if pd.isna(v) or v == 0:
        return "-"
    return round(v, 6)

# ─────────────────────────────────────────────────────────────────────────────
# 7. PROCESSAMENTO DA TABELA
# ─────────────────────────────────────────────────────────────────────────────
if st.session_state['df_bruto'] is not None:
    try:
        # --- AJUSTES MANUAIS ---
        st.write("### 🛠️ Ajustes de Boleta")
        with st.expander("Expandir painel de ajustes (Individual ou Lote)"):
            tab_manual, tab_lote = st.tabs(["Edição Individual", "Upload em Lote"])

            with tab_manual:
                c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
                edit_bol  = c1.text_input("ID Boleta",           key="input_bol")
                edit_vend = c2.text_input("Novo Vendedor",       key="input_vend")
                edit_comp = c3.text_input("Novo Comprador",      key="input_comp")
                edit_cliq = c4.text_input("Novo Cliq Paradigma", key="input_cliq")
                if st.button("Gravar Alteração"):
                    if edit_bol:
                        st.session_state['ajustes_manuais'][tratar_chave(edit_bol)] = {
                            'Vendedor': edit_vend if edit_vend else None,
                            'Comprador': edit_comp if edit_comp else None,
                            'CliqCCEE Paradigma': edit_cliq if edit_cliq else None
                        }
                        salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
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
                            salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
                            st.success("Ajustes em lote carregados!")
                            st.rerun()
                        else:
                            st.error("Coluna 'BOLETA' não encontrada.")
                    except Exception as e_lote:
                        st.error(f"Erro ao processar lote: {e_lote}")

            if st.session_state['ajustes_manuais']:
                st.markdown("---")
                st.write("**Ajustes Ativos:**")
                for bol_id, dados in list(st.session_state['ajustes_manuais'].items()):
                    col_info, col_del = st.columns([6, 1])
                    info_parts = [f"ID: {bol_id}"]
                    if dados['Vendedor']:           info_parts.append(f"Vend: {dados['Vendedor']}")
                    if dados['Comprador']:          info_parts.append(f"Comp: {dados['Comprador']}")
                    if dados['CliqCCEE Paradigma']: info_parts.append(f"Cliq: {dados['CliqCCEE Paradigma']}")
                    col_info.info(" | ".join(info_parts))
                    if col_del.button("Remover", key=f"del_{bol_id}"):
                        del st.session_state['ajustes_manuais'][bol_id]
                        salvar_disco('ajustes_manuais', st.session_state['ajustes_manuais'])
                        st.rerun()
                if st.button("Limpar todos os ajustes"):
                    st.session_state['ajustes_manuais'] = {}
                    salvar_disco('ajustes_manuais', {})
                    st.rerun()

        # --- FILTRAGEM POR MÊS ---
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        mes_num_sel = meses_nomes.index(mes_sel) + 1
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta     = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            df_conferencia['Operacao']     = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte']        = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()

            mapa_energia = {
                'Incentivada-50%':   'Incentivada-I5',
                'Incentivada-100%':  'Incentivada-I1',
                'Incentivada-0%':    'Incentivada-I0',
                'Incentivada-CQ50%': 'Incentivada-CQ5'
            }
            df_conferencia['Tipo Energia'] = (
                df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]])
                .astype(str).str.strip().replace(mapa_energia)
            )

            df_conferencia['Contraparte']      = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP']            = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado']       = (
                df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]])
                .replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})
            )

            df_conferencia['Montante MWh'] = (
                pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
                .fillna(0).round(3)
            )
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)

            df_conferencia['Situacao ERP']       = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC']      = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)

            df_conferencia['% Modulacao Min'] = df_conferencia[col_boleta].map(
                df_lookup[df_base.columns[28]]
            ).apply(tratar_modulacao_pct)
            df_conferencia['% Modulacao Max'] = df_conferencia[col_boleta].map(
                df_lookup[df_base.columns[29]]
            ).apply(tratar_modulacao_pct)

            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            # Ajustes manuais
            df_conferencia['Editado'] = False
            for bol, info in st.session_state['ajustes_manuais'].items():
                mask = df_conferencia['Boleta_Key'] == bol
                if mask.any():
                    df_conferencia.loc[mask, 'Editado'] = True
                    if info['Vendedor']:           df_conferencia.loc[mask, 'Vendedor'] = info['Vendedor']
                    if info['Comprador']:          df_conferencia.loc[mask, 'Comprador'] = info['Comprador']
                    if info['CliqCCEE Paradigma']: df_conferencia.loc[mask, 'CliqCCEE Paradigma'] = info['CliqCCEE Paradigma']

            def resolver_cliq(row):
                vend = row['Vendedor']  if row['Vendedor']  != "-" else ""
                comp = row['Comprador'] if row['Comprador'] != "-" else ""
                if 'BISMUT' in str(row['Parte']).upper():
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'],
                                            st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear', 'db_ccear'), ('cbr', 'db_cbr'), ('matrix', 'db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'],
                                           st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)
            df_conferencia['Modulação CCEE']    = df_conferencia.apply(buscar_modulacao_cliq, axis=1)

            def check_modulacao(row):
                wbc  = str(row['Modulacao WBC']).strip().lower()
                ccee = str(row['Modulação CCEE']).strip().lower()
                if wbc in ['-', '', 'nan'] or ccee in ['-', '', 'nan', 'verificar']: return "-"
                return "OK" if wbc == ccee else "Verificar"

            df_conferencia['Check Modulação'] = df_conferencia.apply(check_modulacao, axis=1)

            df_conferencia['Lim Min CCEE'] = df_conferencia['Contrato CliqCCEE'].apply(
                lambda cod: buscar_limite_cliq(cod, 'LIMITE_MINIMO_MODULACAO_MW'))
            df_conferencia['Lim Max CCEE'] = df_conferencia['Contrato CliqCCEE'].apply(
                lambda cod: buscar_limite_cliq(cod, 'LIMITE_MAXIMO_MODULACAO_MW'))

            def validar_lim_min(row):
                wbc, ccee = row['% Modulacao Min'], row['Lim Min CCEE']
                if str(wbc) in ['-', '', 'nan'] or str(ccee) in ['-', '', 'nan']: return "-"
                try: return "OK" if round(float(wbc), 4) == round(float(ccee), 4) else "Verificar"
                except: return "-"

            def validar_lim_max(row):
                wbc, ccee = row['% Modulacao Max'], row['Lim Max CCEE']
                if str(wbc) in ['-', '', 'nan'] or str(ccee) in ['-', '', 'nan']: return "-"
                try: return "OK" if round(float(wbc), 4) == round(float(ccee), 4) else "Verificar"
                except: return "-"

            df_conferencia['Check Lim Min'] = df_conferencia.apply(validar_lim_min, axis=1)
            df_conferencia['Check Lim Max'] = df_conferencia.apply(validar_lim_max, axis=1)

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

            df_soma_cliq   = df_conferencia[~df_conferencia['Contrato CliqCCEE'].isin(['Verificar', '-', ''])].copy()
            dict_soma_book = df_soma_cliq.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['Volume BOOK'] = df_conferencia['Contrato CliqCCEE'].map(dict_soma_book).fillna(0.0).round(6)

            def buscar_volume_cliq(row):
                cod = row['Contrato CliqCCEE']
                if cod in ['Verificar', '-', '']: return 0.0
                h_mes_valor = h_mes.iloc[0] if not pd.isna(h_mes.iloc[0]) else 744
                for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
                    df_cliq = st.session_state.get(db_key)
                    if df_cliq is not None and cod in df_cliq.index:
                        val = df_cliq.loc[cod, ('MONTANTE_MENSAL_MWh' if cod in CONTRATOS_ESPECIAIS_CCEAR else 'MWmedio')]
                        val = val.iloc[0] if isinstance(val, pd.Series) else val
                        if not pd.isna(val) and val != "":
                            v = float(str(val).replace(',', '.'))
                            return v / h_mes_valor if cod in CONTRATOS_ESPECIAIS_CCEAR else v
                return 0.0

            df_conferencia['Volume CliqCCEE'] = df_conferencia.apply(buscar_volume_cliq, axis=1).fillna(0.0).round(6)

            def validar_volume_logic(row):
                if row['Contrato CliqCCEE'] in ['Verificar', '-', '']: return "-"
                return "OK" if round(row['Volume BOOK'], 6) == round(row['Volume CliqCCEE'], 6) else "VERIFICAR"

            df_conferencia['Validação Volume'] = df_conferencia.apply(validar_volume_logic, axis=1)

            df_pagos = df_conferencia[df_conferencia['Situacao ERP'].astype(str).str.upper() == 'PAGO'].copy()
            dict_soma_pagos = df_pagos.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()

            def validar_pagamento(row):
                if row['Contrato CliqCCEE'] in ['Verificar', '-', '']: return "-"
                total_pago = dict_soma_pagos.get(row['Contrato CliqCCEE'], 0.0)
                return "Pago" if round(total_pago, 6) >= round(row['Volume BOOK'], 6) and row['Volume BOOK'] > 0 else "-"

            df_conferencia['SITUAÇÃO PGTO']        = df_conferencia.apply(validar_pagamento, axis=1)
            df_conferencia['Pendência Financeira'] = (
                df_conferencia['Razao Social'].str.strip().str.upper()
                .map(st.session_state['dict_pendencias']).fillna(0.0)
            )

            df_conferencia['Varejista'] = df_conferencia['Comprador'].apply(classificar_varejista)

            # --- FILTROS ---
            st.write("### Filtros")
            f1, f2, f3, f4, f5 = st.columns([1.5, 1.5, 1.5, 1.5, 1.5])
            op_f        = f1.selectbox("Operação",          ["Todos"] + sorted(df_conferencia['Operacao'].unique()))
            parte_f     = f2.selectbox("Parte",             ["Todos"] + sorted(df_conferencia['Parte'].unique()))
            cliq_f      = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].unique()))
            valid_vol_f = f4.selectbox("Validação Volume",  ["Todos"] + sorted(df_conferencia['Validação Volume'].unique()))

            todas_opts_mod = sorted(set(
                list(df_conferencia['Check Modulação'].unique()) +
                list(df_conferencia['Check Lim Min'].unique()) +
                list(df_conferencia['Check Lim Max'].unique())
            ))
            check_mod_f = f5.selectbox("Check Modulação / Limites", ["Todos"] + todas_opts_mod)

            fa1, fa2 = st.columns([3, 1])

            opcoes_contraparte = sorted(
                [str(x) for x in df_conferencia['Contraparte'].dropna().unique() if str(x).strip() != ""]
            )
            contraparte_f = fa1.multiselect(
                "Contraparte (selecione uma ou mais — vazio = todas)",
                options=opcoes_contraparte,
                default=[],
                placeholder="Todas as contrapartes"
            )

            varejista_f = fa2.selectbox(
                "Varejista",
                options=["Todos", "Sim", "Não"],
                key="filtro_varejista"
            )

            f6, f7, f8 = st.columns([1.5, 1.5, 1.5])
            zerar_intra   = f6.toggle("Zerar Intraportfólio")
            zerar_entre   = f7.toggle("Zerar Entre Empresas")
            ocultar_vazio = f8.toggle("Ocultar Volumes Zerados")

            df_final = df_conferencia.copy()
            if op_f        != "Todos": df_final = df_final[df_final['Operacao'] == op_f]
            if parte_f     != "Todos": df_final = df_final[df_final['Parte'] == parte_f]
            if cliq_f      != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]
            if valid_vol_f != "Todos": df_final = df_final[df_final['Validação Volume'] == valid_vol_f]
            if check_mod_f != "Todos":
                mask_mod = (
                    (df_final['Check Modulação'] == check_mod_f) |
                    (df_final['Check Lim Min']   == check_mod_f) |
                    (df_final['Check Lim Max']   == check_mod_f)
                )
                df_final = df_final[mask_mod]
            if contraparte_f:
                df_final = df_final[df_final['Contraparte'].astype(str).isin(contraparte_f)]
            if varejista_f != "Todos":
                df_final = df_final[df_final['Varejista'] == varejista_f]

            if zerar_intra:
                mask_i = df_final['Vendedor'].str.lower().str.strip() == df_final['Comprador'].str.lower().str.strip()
                df_final.loc[mask_i, ['Montante MWh', 'Volume MWm']] = 0.0
            if zerar_entre:
                mask_p = df_final['Parte'].str.contains("BISMUT|GET", na=False, case=False)
                mask_c = (df_final['Contraparte'].str.upper().str.startswith("MATRIX", na=False) &
                          ~df_final['Contraparte'].str.upper().str.contains("MATRIX VAR", na=False))
                df_final.loc[mask_p & mask_c, ['Montante MWh', 'Volume MWm']] = 0.0
            if ocultar_vazio:
                df_final = df_final[df_final['Volume MWm'] != 0]

            # --- RESUMO DE OPERAÇÕES ---
            st.write("### 📦 Resumo de Operações")
            n_compras = int((df_final['Operacao'].astype(str).str.upper() == 'COMPRA').sum())
            n_vendas  = int((df_final['Operacao'].astype(str).str.upper() == 'VENDA').sum())
            n_total   = len(df_final)
            n_outros  = n_total - n_compras - n_vendas

            mc1, mc2, mc3, mc4 = st.columns(4)
            mc1.metric("🛒 Compras", n_compras)
            mc2.metric("📤 Vendas",  n_vendas)
            mc3.metric("🔄 Outros",  n_outros)
            mc4.metric("📋 Total",   n_total)

            # --- VALIDAÇÃO DE MATCH CCEE ---
            bases_carregadas = [
                k.replace('db_', '').upper()
                for k in ['db_ccear', 'db_cbr', 'db_matrix', 'db_bismut']
                if st.session_state.get(k) is not None
            ]
            with st.expander("🔍 Validação de Match CCEE", expanded=False):
                if not bases_carregadas:
                    st.warning("Nenhuma base Cliq CCEE carregada.")
                else:
                    _, sem_match, _ = gerar_relatorio_match(df_final)
                    if sem_match.empty:
                        st.success("Nenhuma linha sem match!")
                    else:
                        st.warning(f"{len(sem_match)} linha(s) sem contrato correspondente.")
                        st.dataframe(sem_match.reset_index(drop=True), use_container_width=True, hide_index=True)

            ordem = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 'CP/LP',
                'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm',
                'CliqCCEE Paradigma',
                'Modulacao WBC', 'Modulação CCEE', 'Check Modulação',
                '% Modulacao Min', 'Lim Min CCEE', 'Check Lim Min',
                '% Modulacao Max', 'Lim Max CCEE', 'Check Lim Max',
                'Contrato CliqCCEE mes anterior', 'Vendedor', 'Comprador', 'Varejista',
                'Contrato CliqCCEE',
                'Status do Contrato', 'SITUAÇÃO PGTO', 'Volume BOOK', 'Volume CliqCCEE',
                'Validação Volume', 'Situacao ERP', 'Razao Social', 'Pendência Financeira', 'Editado'
            ]

            def highlight_rows(row):
                return ['background-color: #fff4cc'] * len(row) if row['Editado'] else [''] * len(row)

            col_config_main = {
                "Editado":              None,
                "Montante MWh":         st.column_config.NumberColumn(format="%.3f"),
                "Volume MWm":           st.column_config.NumberColumn(format="%.6f"),
                "Volume BOOK":          st.column_config.NumberColumn(format="%.6f"),
                "Volume CliqCCEE":      st.column_config.NumberColumn(format="%.6f"),
                "Lim Min CCEE":         st.column_config.NumberColumn(format="%.6f"),
                "Lim Max CCEE":         st.column_config.NumberColumn(format="%.6f"),
                "Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f"),
                "Varejista":            st.column_config.TextColumn("Varejista"),
                "% Modulacao Min":      st.column_config.TextColumn("% Mod Min"),
                "% Modulacao Max":      st.column_config.TextColumn("% Mod Max"),
            }

            st.dataframe(
                df_final[ordem].sort_values(by=col_boleta).style.apply(highlight_rows, axis=1),
                use_container_width=True,
                hide_index=True,
                column_config=col_config_main,
            )

            # ═════════════════════════════════════════════════════════════════
            # HELPER: filtra df_conferencia por Parte e aplica flags de zeragem
            # ═════════════════════════════════════════════════════════════════
            def df_wbc_para_parte(nome_parte_upper):
                df_w = df_conferencia[
                    df_conferencia['Parte'].astype(str).str.strip().str.upper() == nome_parte_upper
                ].copy()
                if zerar_intra:
                    mask_zi = df_w['Vendedor'].str.lower().str.strip() == df_w['Comprador'].str.lower().str.strip()
                    df_w.loc[mask_zi, 'Volume MWm'] = 0.0
                if zerar_entre:
                    mask_zp = df_w['Parte'].str.contains("BISMUT|GET", na=False, case=False)
                    mask_zc = (df_w['Contraparte'].str.upper().str.startswith("MATRIX", na=False) &
                               ~df_w['Contraparte'].str.upper().str.contains("MATRIX VAR", na=False))
                    df_w.loc[mask_zp & mask_zc, 'Volume MWm'] = 0.0
                return df_w

            # ─────────────────────────────────────────────────────────────────
            # BLOCO BISMUT
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — BISMUT CCEE  |  BISMUT WBC")

            db_bismut_ccee = st.session_state.get('db_bismut')

            submercados_bismut = ["Todos"]
            if db_bismut_ccee is not None:
                df_bism_temp = db_bismut_ccee.reset_index()
                if 'SUBMERCADO_ENTREGA' in df_bism_temp.columns:
                    subs_unicos = sorted(
                        df_bism_temp['SUBMERCADO_ENTREGA'].dropna().astype(str).str.strip()
                        .replace('', pd.NA).dropna().unique().tolist()
                    )
                    submercados_bismut += subs_unicos

            filtro_sub_bismut = st.selectbox(
                "Filtro Submercado (Tabela CCEE Bismut)",
                options=submercados_bismut,
                key="filtro_submercado_bismut_ccee"
            )

            PERFIS_BISMUT = [
                "BISMUT COM I5",
                "BISMUT COM I0",
                "BISMUT COM I1",
                "BISMUT COM",
            ]
            PARTE_BISMUT_WBC = "BISMUT COMERCIALIZADORA DE ENERGIA S/A"

            df_bismut_ccee_tab = build_ccee_tabela(PERFIS_BISMUT, ['db_bismut'], filtro_sub_bismut)
            df_bismut_wbc_tab  = build_wbc_tabela(PERFIS_BISMUT, df_wbc_para_parte(PARTE_BISMUT_WBC.upper()))

            render_tabela_par(
                "BISMUT CCEE", f"BISMUT WBC — {PARTE_BISMUT_WBC}",
                df_bismut_ccee_tab, df_bismut_wbc_tab,
                aviso_sem_base="Base Cliq Bismut não carregada." if db_bismut_ccee is None else None
            )
            render_reconciliacao("BISMUT", df_bismut_ccee_tab, df_bismut_wbc_tab)

            # ─────────────────────────────────────────────────────────────────
            # BLOCO MATRIX
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — MATRIX CCEE  |  MATRIX WBC")

            DBS_MATRIX = ['db_ccear', 'db_cbr', 'db_matrix']

            submercados_matrix = ["Todos"]
            for db_key in DBS_MATRIX:
                df_m = st.session_state.get(db_key)
                if df_m is not None:
                    df_mt = df_m.reset_index()
                    if 'SUBMERCADO_ENTREGA' in df_mt.columns:
                        for s in sorted(df_mt['SUBMERCADO_ENTREGA'].dropna().astype(str).str.strip()
                                        .replace('', pd.NA).dropna().unique().tolist()):
                            if s not in submercados_matrix:
                                submercados_matrix.append(s)
            submercados_matrix = ["Todos"] + sorted(submercados_matrix[1:])

            filtro_sub_matrix = st.selectbox(
                "Filtro Submercado (Tabela Matrix CCEE)",
                options=submercados_matrix,
                key="filtro_submercado_matrix_ccee"
            )

            PERFIS_MATRIX = [
                "MATRIX COM I5",
                "MATRIX COM",
                "MATRIX COM I1",
                "MATRIX COM I0",
                "MATRIX COM CQ5",
            ]
            PARTE_MATRIX_WBC = "MATRIX COMERCIALIZADORA DE ENERGIA ELETRICA S/A"

            df_matrix_ccee_tab = build_ccee_tabela(PERFIS_MATRIX, DBS_MATRIX, filtro_sub_matrix)
            df_matrix_wbc_tab  = build_wbc_tabela(PERFIS_MATRIX, df_wbc_para_parte(PARTE_MATRIX_WBC.upper()))

            bases_matrix_ok = any(st.session_state.get(k) is not None for k in DBS_MATRIX)
            render_tabela_par(
                "MATRIX CCEE", f"MATRIX WBC — {PARTE_MATRIX_WBC}",
                df_matrix_ccee_tab, df_matrix_wbc_tab,
                aviso_sem_base="Nenhuma base Cliq Matrix/CCEAR/CBR carregada." if not bases_matrix_ok else None
            )
            render_reconciliacao("MATRIX", df_matrix_ccee_tab, df_matrix_wbc_tab)

            # ─────────────────────────────────────────────────────────────────
            # BLOCO GET ENERGY TRADING
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — GET ENERGY TRADING CCEE  |  GET ENERGY TRADING WBC")

            filtro_sub_get = st.selectbox(
                "Filtro Submercado (Tabela GET CCEE)",
                options=submercados_bismut,
                key="filtro_submercado_get_ccee"
            )

            PERFIS_GET = [
                "GET ENERGY TRADING",
                "GET ENERGY TRADING I5",
                "GET ENERGY TRADING I0",
                "GET ENERGY TRADING I1",
                "GET ENERGY TRADING CQ5",
            ]
            PARTE_GET_WBC = "GET ENERGY TRADING"

            df_get_ccee_tab = build_ccee_tabela(PERFIS_GET, ['db_bismut'], filtro_sub_get)
            df_get_wbc_tab  = build_wbc_tabela(PERFIS_GET, df_wbc_para_parte(PARTE_GET_WBC.upper()))

            render_tabela_par(
                "GET ENERGY TRADING CCEE", f"GET ENERGY TRADING WBC — {PARTE_GET_WBC}",
                df_get_ccee_tab, df_get_wbc_tab,
                aviso_sem_base="Base Cliq Bismut não carregada." if db_bismut_ccee is None else None
            )
            render_reconciliacao("GET ENERGY TRADING", df_get_ccee_tab, df_get_wbc_tab)

            # ─────────────────────────────────────────────────────────────────
            # BLOCO CINERGY
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — CINERGY CCEE  |  CINERGY WBC")

            filtro_sub_cinergy = st.selectbox(
                "Filtro Submercado (Tabela CINERGY CCEE)",
                options=submercados_bismut,
                key="filtro_submercado_cinergy_ccee"
            )

            PERFIS_CINERGY = [
                "CINERGY COM",
                "CINERGY COM I1",
                "CINERGY COM I5",
                "CINERGY COM I0",
                "CINERGY COM I8",
                "CINERGY COM I5 2",
                "CINERGY COM I1 2",
                "CINERGY COM I0 2",
                "CINERGY COM I8 2",
            ]
            PARTE_CINERGY_WBC = "CINERGY COM"

            df_cinergy_ccee_tab = build_ccee_tabela(PERFIS_CINERGY, ['db_bismut'], filtro_sub_cinergy)
            df_cinergy_wbc_tab  = build_wbc_tabela(PERFIS_CINERGY, df_wbc_para_parte(PARTE_CINERGY_WBC.upper()))

            render_tabela_par(
                "CINERGY CCEE", f"CINERGY WBC — {PARTE_CINERGY_WBC}",
                df_cinergy_ccee_tab, df_cinergy_wbc_tab,
                aviso_sem_base="Base Cliq Bismut não carregada." if db_bismut_ccee is None else None
            )
            render_reconciliacao("CINERGY", df_cinergy_ccee_tab, df_cinergy_wbc_tab)

            # ─────────────────────────────────────────────────────────────────
            # BLOCO MTX CAMANDUCAIA
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — MTX CAMANDUCAIA CCEE  |  MTX CAMANDUCAIA WBC")

            filtro_sub_mtx = st.selectbox(
                "Filtro Submercado (Tabela MTX CAMANDUCAIA CCEE)",
                options=submercados_bismut,
                key="filtro_submercado_mtx_ccee"
            )

            PERFIS_MTX = [
                "MTX CAMANDUCAIA",
            ]
            PARTE_MTX_WBC = "MTX CAMANDUCAIA"

            df_mtx_ccee_tab = build_ccee_tabela(PERFIS_MTX, ['db_bismut'], filtro_sub_mtx)
            df_mtx_wbc_tab  = build_wbc_tabela(PERFIS_MTX, df_wbc_para_parte(PARTE_MTX_WBC.upper()))

            render_tabela_par(
                "MTX CAMANDUCAIA CCEE", f"MTX CAMANDUCAIA WBC — {PARTE_MTX_WBC}",
                df_mtx_ccee_tab, df_mtx_wbc_tab,
                aviso_sem_base="Base Cliq Bismut não carregada." if db_bismut_ccee is None else None
            )
            render_reconciliacao("MTX CAMANDUCAIA", df_mtx_ccee_tab, df_mtx_wbc_tab)

            # ─────────────────────────────────────────────────────────────────
            # BLOCO ARGENTUM
            # ─────────────────────────────────────────────────────────────────
            st.write("### 📊 Tabela — ARGENTUM CCEE  |  ARGENTUM WBC")

            filtro_sub_argentum = st.selectbox(
                "Filtro Submercado (Tabela ARGENTUM CCEE)",
                options=submercados_bismut,
                key="filtro_submercado_argentum_ccee"
            )

            PERFIS_ARGENTUM = [
                "ARGENTUM COM",
                "ARGENTUM COM I1",
                "ARGENTUM COM I5",
                "ARGENTUM COM I0",
                "ARGENTUM COM I8",
            ]
            PARTE_ARGENTUM_WBC = "ARGENTUM COM"

            df_argentum_ccee_tab = build_ccee_tabela(PERFIS_ARGENTUM, ['db_bismut'], filtro_sub_argentum)
            df_argentum_wbc_tab  = build_wbc_tabela(PERFIS_ARGENTUM, df_wbc_para_parte(PARTE_ARGENTUM_WBC.upper()))

            render_tabela_par(
                "ARGENTUM CCEE", f"ARGENTUM WBC — {PARTE_ARGENTUM_WBC}",
                df_argentum_ccee_tab, df_argentum_wbc_tab,
                aviso_sem_base="Base Cliq Bismut não carregada." if db_bismut_ccee is None else None
            )
            render_reconciliacao("ARGENTUM", df_argentum_ccee_tab, df_argentum_wbc_tab)

        else:
            st.warning("Sem dados para este período.")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
