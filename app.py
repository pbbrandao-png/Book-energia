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
    zip_id = (zip_file_obj.name, zip_file_obj.size)
    key_id    = f"{state_prefix}_zip_id"
    key_pasta = f"{state_prefix}_zip_pasta"
    key_sel   = f"{state_prefix}_zip_selected"
    key_bytes = f"{state_prefix}_zip_bytes"

    if st.session_state.get(key_id) != zip_id:
        st.session_state[key_id]    = zip_id
        st.session_state[key_pasta] = ''
        st.session_state[key_sel]   = None
        st.session_state[key_bytes] = zip_file_obj.read()

    zf = zipfile.ZipFile(io.BytesIO(st.session_state[key_bytes]))
    pasta_atual = st.session_state.get(key_pasta, '')
    subpastas, arquivos = _listar_zip(zf, pasta_atual)

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

    for pasta in subpastas:
        nome_exib = pasta[len(pasta_atual):].strip('/')
        if container.button(f"📁 {nome_exib}", key=f"{state_prefix}_pasta_{pasta}"):
            st.session_state[key_pasta] = pasta
            st.session_state[key_sel]   = None
            st.rerun()

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
    tipos_aceitos = list(set(types + ['zip']))
    arq = container.file_uploader(label, type=tipos_aceitos, key=key)

    if arq is None:
        return None

    if arq.name.lower().endswith('.zip'):
        with container.expander("📦 Navegar dentro do ZIP", expanded=True):
            return zip_browser_inline(arq, state_prefix, st.sidebar if container is st.sidebar else st)
    else:
        return arq


# ─────────────────────────────────────────────────────────────────────────────
# ★ ZIP MULTI-BASE — autodetecção por padrão de nome de arquivo
# ─────────────────────────────────────────────────────────────────────────────

def _listar_planilhas_zip(zf):
    """Retorna lista de todos os caminhos de planilha dentro do ZIP."""
    return sorted(
        nome for nome in zf.namelist()
        if nome.lower().endswith(EXTS_PLANILHA) and not nome.startswith('__MACOSX')
    )


def detectar_base_por_nome(nome_arquivo):
    """
    Mapeia nome de arquivo para chave de base pelo padrão de nome.
    Regras:
      ccear_q_XXXXXX.* → db_ccear
      cbr_mercado_proprio_XXXX (sem 'parcela') → db_cbr
      cceal_firme_XXXXXX.* (sem 'parcela') → db_bismut
    """
    n = nome_arquivo.lower()
    if n.startswith('ccear_q_'):
        return 'db_ccear'
    if n.startswith('cbr_mercado_proprio_') and 'parcela' not in n:
        return 'db_cbr'
    if n.startswith('cceal_firme_') and 'parcela' not in n:
        return 'db_bismut'
    return None


def zip_multi_base_sidebar():
    st.sidebar.markdown("---")
    st.sidebar.subheader("📦 Bases Cliq CCEE")

    resultado = {k: None for k in ['db_ccear', 'db_cbr', 'db_matrix', 'db_bismut']}

    BASES_LABELS = {
        'db_ccear':  'CCEAR_Q',
        'db_cbr':    'CBR Mercado',
        'db_matrix': 'Matrix',
        'db_bismut': 'Bismut (cceal_firme)',
    }

    def _autodetectar_zip(zip_obj, prefix, bases_alvo, container):
        """
        Lê o ZIP, detecta automaticamente os arquivos pelas regras de nome
        e retorna {db_key: buf}. Para arquivos não reconhecidos mostra
        selectbox manual apenas para as bases que faltaram.
        """
        parcial = {}
        zip_id    = (zip_obj.name, zip_obj.size)
        key_id    = f"{prefix}_id"
        key_bytes = f"{prefix}_bytes"

        if st.session_state.get(key_id) != zip_id:
            st.session_state[key_id]    = zip_id
            st.session_state[key_bytes] = zip_obj.read()
            # Limpa seleções manuais anteriores ao trocar o ZIP
            for base_k in BASES_LABELS:
                st.session_state.pop(f"{prefix}_manual_{base_k}", None)

        raw_bytes = st.session_state[key_bytes]
        zf = zipfile.ZipFile(io.BytesIO(raw_bytes))
        planilhas = _listar_planilhas_zip(zf)

        # ── Autodetecção ──────────────────────────────────────────────────
        detectados = {}
        for caminho in planilhas:
            nome_curto = caminho.split('/')[-1]
            base_k = detectar_base_por_nome(nome_curto)
            if base_k and base_k in bases_alvo and base_k not in detectados:
                detectados[base_k] = caminho

        if detectados:
            container.caption(f"📦 `{zip_obj.name}` — detectado automaticamente:")
            for base_k, caminho in detectados.items():
                nome_curto = caminho.split('/')[-1]
                container.success(f"✅ **{BASES_LABELS[base_k]}** → `{nome_curto}`")
                data = zf.read(caminho)
                buf  = io.BytesIO(data)
                buf.name = nome_curto
                buf.size = len(data)
                parcial[base_k] = buf
        else:
            container.warning(f"Nenhum arquivo reconhecido automaticamente em `{zip_obj.name}`.")

        # ── Selectbox manual apenas para bases que não foram detectadas ───
        bases_faltando = [k for k in bases_alvo if k not in detectados]
        if bases_faltando and planilhas:
            opcoes_manual = ["— não usar —"] + [
                p.split('/')[-1] + f"  ({p})" for p in planilhas
            ]
            mapa_opcoes = {opcoes_manual[i + 1]: planilhas[i] for i in range(len(planilhas))}
            container.caption("Atribuição manual para arquivos não reconhecidos:")
            for base_k in bases_faltando:
                sel_key = f"{prefix}_manual_{base_k}"
                escolha = container.selectbox(
                    f"Arquivo para **{BASES_LABELS[base_k]}**",
                    options=opcoes_manual,
                    index=0,
                    key=sel_key,
                )
                if escolha != "— não usar —" and escolha in mapa_opcoes:
                    caminho = mapa_opcoes[escolha]
                    data = zf.read(caminho)
                    buf  = io.BytesIO(data)
                    buf.name = caminho.split('/')[-1]
                    buf.size = len(data)
                    parcial[base_k] = buf

        return parcial

    # ── ZIP 1: CBR / CCEAR (+ Matrix se houver) ───────────────────────────
    st.sidebar.markdown(
        f"{status_icon('db_ccear')} {status_icon('db_cbr')} {status_icon('db_matrix')} "
        f"**ZIP 1 — CBR / CCEAR / Matrix**"
    )
    zip1 = st.sidebar.file_uploader(
        "Subir ZIP (CBR / CCEAR / Matrix)", type=['zip', 'xlsx', 'xlsm', 'csv'],
        key="up_zip1_cliq"
    )

    if zip1 is not None:
        if zip1.name.lower().endswith('.zip'):
            with st.sidebar.expander("⚙️ ZIP 1 — detalhes", expanded=True):
                parcial1 = _autodetectar_zip(
                    zip1, "zip1", ['db_ccear', 'db_cbr', 'db_matrix'], st.sidebar
                )
            for k, v in parcial1.items():
                if v is not None:
                    resultado[k] = v
        else:
            # Arquivo avulso: tenta autodetectar pelo nome
            base_auto = detectar_base_por_nome(zip1.name)
            if base_auto:
                st.sidebar.success(f"✅ Detectado como **{BASES_LABELS[base_auto]}**")
                resultado[base_auto] = zip1
            else:
                labels_zip1 = {k: v for k, v in BASES_LABELS.items() if k != 'db_bismut'}
                base_direta = st.sidebar.selectbox(
                    f"'{zip1.name}' é qual base?",
                    options=list(labels_zip1.values()),
                    key="up_zip1_base_direta"
                )
                chave_direta = [k for k, v in labels_zip1.items() if v == base_direta][0]
                resultado[chave_direta] = zip1

    # ── ZIP 2: Bismut ──────────────────────────────────────────────────────
    st.sidebar.markdown(f"{status_icon('db_bismut')} **ZIP 2 — Bismut**")
    zip2 = st.sidebar.file_uploader(
        "Subir ZIP (Bismut)", type=['zip', 'xlsx', 'xlsm', 'csv'],
        key="up_zip2_cliq"
    )

    if zip2 is not None:
        if zip2.name.lower().endswith('.zip'):
            with st.sidebar.expander("⚙️ ZIP 2 — detalhes", expanded=True):
                parcial2 = _autodetectar_zip(zip2, "zip2", ['db_bismut'], st.sidebar)
            for k, v in parcial2.items():
                if v is not None:
                    resultado[k] = v
        else:
            base_auto2 = detectar_base_por_nome(zip2.name)
            if base_auto2 == 'db_bismut':
                st.sidebar.success("✅ Detectado como **Bismut**")
                resultado['db_bismut'] = zip2
            else:
                resultado['db_bismut'] = zip2

    return resultado


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
    
    # Lógica atualizada: apenas Bismut usa db_bismut. Cinergy e outras usam Matrix/CCEAR/CBR.
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
        
        # Cinergy, Get, Argentum e Camanducaia foram removidas da base Bismut aqui
        is_bismut  = any(p in str(row.get('Parte', '')).upper() for p in ('BISMUT',))
        
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

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'fid_cceal2',
              'fid_mapa', 'fid_pendencias', 'fid_cliq_multi']:
    if chave not in st.session_state:
        st.session_state[chave] = None

# ─────────────────────────────────────────────────────────────────────────────
# 5. SIDEBAR — arquivos gerais
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

# ── Uploader unificado para bases Cliq ───────────────────────────────────────
arquivos_cliq = zip_multi_base_sidebar()

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

# ── Carregamento das bases Cliq ───────────────────────────────────────────────
fid_cliq_atual = tuple(get_file_id(arquivos_cliq[k]) for k in ['db_ccear', 'db_cbr', 'db_matrix', 'db_bismut'])
if fid_cliq_atual != st.session_state.get('fid_cliq_multi'):
    st.session_state['fid_cliq_multi'] = fid_cliq_atual
    for db_key in ['db_ccear', 'db_cbr', 'db_matrix', 'db_bismut']:
        arq = arquivos_cliq[db_key]
        if arq is not None:
            val = carregar_csv_cliq(arq)
            st.session_state[db_key] = val
            salvar_disco(db_key, val)

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
        perfil_upper = perfil.strip().upper()
        mask_c = df_wbc_base['Comprador'].astype(str).str.strip().str.upper() == perfil_upper
        mask_v = df_wbc_base['Vendedor'].astype(str).str.strip().str.upper()  == perfil_upper
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
                use_container_width=True, hide_index=True, column_config=COL_CONFIG_PERFIL
            )
    with col_r:
        st.markdown(f"**{titulo_wbc}**")
        st.dataframe(
            df_wbc.style.apply(highlight_total, axis=1),
            use_container_width=True, hide_index=True, column_config=COL_CONFIG_PERFIL
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
        diff = row['DIFERENÇA']
        if abs(diff) > 0:
            causas.append(
                f"**{perfil}**: NET CCEE = {row['NET_CCEE']:.6f} | NET WBC = {row['NET_WBC']:.6f} | "
                f"Diferença = **{diff:.6f}**"
            )

    with st.expander(f"🔎 Reconciliação NET — {titulo}", expanded=False):
        st.dataframe(
            df_rec.style.apply(
                lambda r: ['background-color: #fdecea' if r['STATUS'] != '✅ OK' else '' for _ in r], axis=1
            ),
            use_container_width=True, hide_index=True,
            column_config={
                "PERFIL": st.column_config.TextColumn("PERFIL"),
                "NET_CCEE": st.column_config.NumberColumn("NET CCEE", format="%.6f"),
                "NET_WBC": st.column_config.NumberColumn("NET WBC", format="%.6f"),
                "DIFERENÇA": st.column_config.NumberColumn("Diferença", format="%.6f"),
                "STATUS": st.column_config.TextColumn("Status"),
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
# 7. RENDERIZAÇÃO DAS TABELAS NO FINAL DO SCRIPT (EXEMPLO CINERGY NA MATRIX)
# ─────────────────────────────────────────────────────────────────────────────

if st.session_state['df_bruto'] is not None:
    df_wbc_base = st.session_state['df_bruto'].copy()
    
    # ... (Seu código de processamento da tabela principal aqui)

    # BLOCO CINERGY (AGORA BUSCANDO NA MATRIX)
    st.write("### 📊 Tabela — CINERGY")
    df_cin_ccee_tab = build_ccee_tabela(["CINERGY"], ['db_matrix'], "Todos")
    df_cin_wbc_tab  = build_wbc_tabela(["CINERGY"], df_wbc_base)
    render_tabela_par("CINERGY CCEE", "CINERGY WBC", df_cin_ccee_tab, df_cin_wbc_tab)
    render_reconciliacao("CINERGY", df_cin_ccee_tab, df_cin_wbc_tab)

    # BLOCO GET (MATRIX)
    st.write("### 📊 Tabela — GET")
    df_get_ccee_tab = build_ccee_tabela(["GET"], ['db_matrix'], "Todos")
    df_get_wbc_tab  = build_wbc_tabela(["GET"], df_wbc_base)
    render_tabela_par("GET CCEE", "GET WBC", df_get_ccee_tab, df_get_wbc_tab)
    render_reconciliacao("GET", df_get_ccee_tab, df_get_wbc_tab)

    # BLOCO MTX CAMANDUCAIA (MATRIX)
    st.write("### 📊 Tabela — MTX CAMANDUCAIA")
    df_mtx_ccee_tab = build_ccee_tabela(["MTX CAMANDUCAIA"], ['db_matrix'], "Todos")
    df_mtx_wbc_tab  = build_wbc_tabela(["MTX CAMANDUCAIA"], df_wbc_base)
    render_tabela_par("MTX CAMANDUCAIA CCEE", "MTX CAMANDUCAIA WBC", df_mtx_ccee_tab, df_mtx_wbc_tab)
    render_reconciliacao("MTX CAMANDUCAIA", df_mtx_ccee_tab, df_mtx_wbc_tab)

    # BLOCO ARGENTUM (MATRIX)
    st.write("### 📊 Tabela — ARGENTUM")
    PERFIS_ARGENTUM = ["ARGENTUM COM", "ARGENTUM COM I1", "ARGENTUM COM I5", "ARGENTUM COM I0", "ARGENTUM COM I8"]
    df_arg_ccee_tab = build_ccee_tabela(PERFIS_ARGENTUM, ['db_matrix'], "Todos")
    df_arg_wbc_tab  = build_wbc_tabela(PERFIS_ARGENTUM, df_wbc_base)
    render_tabela_par("ARGENTUM CCEE", "ARGENTUM WBC", df_arg_ccee_tab, df_arg_wbc_tab)
    render_reconciliacao("ARGENTUM", df_arg_ccee_tab, df_arg_wbc_tab)
