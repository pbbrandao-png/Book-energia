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
    """Normaliza string para comparação: strip + lowercase."""
    if pd.isna(valor): return ""
    return str(valor).strip().lower()

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

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq, tipo_base,
                    nome_vendedor, nome_comprador):
    if df_cliq is None:
        return "Verificar"

    mapa     = COLUNAS_CLIQ.get(tipo_base, {})
    col_vend = mapa.get('vendedor')
    col_comp = mapa.get('comprador')

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index:
            return False

        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]

        situacao = str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper()
        if situacao == 'RASCUNHO':
            return False

        if col_vend and col_vend in df_cliq.columns:
            vend_exp  = limpar_str(nome_vendedor)
            vend_cliq = limpar_str(row.get(col_vend, ''))
            if vend_exp and vend_cliq != vend_exp:
                return False

        if col_comp and col_comp in df_cliq.columns:
            comp_exp  = limpar_str(nome_comprador)
            comp_cliq = limpar_str(row.get(col_comp, ''))
            if comp_exp and comp_cliq != comp_exp:
                return False

        return True

    if checar(cod_paradigma):
        return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior):
        return tratar_chave(cod_mes_anterior)
    return "Verificar"

# 3. INICIALIZA SESSION STATE
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
anos = [str(a) for a in range(2024, 2031)]

if 'mes_sel' not in st.session_state: st.session_state['mes_sel'] = meses_nomes[datetime.now().month - 1]
if 'ano_sel' not in st.session_state: st.session_state['ano_sel'] = str(datetime.now().year)

for chave in ['df_bruto', 'dict_mes_anterior', 'dict_comprador', 'dict_vendedor',
              'db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
    if chave not in st.session_state:
        st.session_state[chave] = {} if 'dict' in chave else None

for chave in ['fid_subido', 'fid_anterior', 'fid_pessoas', 'chave_matrix', 'fid_cceal2']:
    if chave not in st.session_state: st.session_state[chave] = None

# 4. INTERFACE LATERAL
st.sidebar.title("Configurações")
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes,
    index=meses_nomes.index(st.session_state['mes_sel']), key='mes_sel')
ano_sel = st.sidebar.selectbox("Ano", anos,
    index=anos.index(st.session_state['ano_sel']) if st.session_state['ano_sel'] in anos else 0,
    key='ano_sel')
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.markdown("---")
def get_file_id(arq): return (arq.name, arq.size) if arq else None

arquivo_subido   = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior.xlsx",      type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. Exportador (4).xlsx",          type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q",     type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado", type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq Matrix",      type=['xlsx', 'csv'])
arq_cceal2 = st.sidebar.file_uploader("Cliq Bismut",      type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

# 5. CARREGAMENTO DOS DADOS
fid_subido = get_file_id(arquivo_subido)
if fid_subido != st.session_state['fid_subido']:
    st.session_state['fid_subido'] = fid_subido
    if arquivo_subido:
        try: st.session_state['df_bruto'] = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        except: st.session_state['df_bruto'] = None

fid_anterior = get_file_id(arquivo_anterior)
if fid_anterior != st.session_state['fid_anterior']:
    st.session_state['fid_anterior'] = fid_anterior
    dict_ma = {}
    if arquivo_anterior:
        try:
            df_apoio = pd.read_excel(arquivo_anterior, dtype=str)
            dict_ma = pd.Series(
                df_apoio.iloc[:, 1].values,
                index=df_apoio.iloc[:, 0].apply(tratar_chave).values
            ).to_dict()
        except: pass
    st.session_state['dict_mes_anterior'] = dict_ma

fid_pessoas = get_file_id(arquivo_pessoas)
if fid_pessoas != st.session_state['fid_pessoas']:
    st.session_state['fid_pessoas'] = fid_pessoas
    dict_v, dict_c = {}, {}
    if arquivo_pessoas:
        try:
            df_pers = pd.read_excel(arquivo_pessoas)
            df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
            dict_c = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
            dict_v = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
        except: pass
    st.session_state['dict_comprador'] = dict_c
    st.session_state['dict_vendedor']  = dict_v

if (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1)) != st.session_state['chave_matrix']:
    st.session_state['chave_matrix'] = (get_file_id(arq_ccear), get_file_id(arq_cbr), get_file_id(arq_cceal1))
    st.session_state['db_ccear']  = carregar_csv_cliq(arq_ccear)
    st.session_state['db_cbr']    = carregar_csv_cliq(arq_cbr)
    st.session_state['db_matrix'] = carregar_csv_cliq(arq_cceal1)

if get_file_id(arq_cceal2) != st.session_state['fid_cceal2']:
    st.session_state['fid_cceal2'] = get_file_id(arq_cceal2)
    st.session_state['db_bismut']  = carregar_csv_cliq(arq_cceal2)

# 6. PROCESSAMENTO
df_bruto = st.session_state['df_bruto']
if df_bruto is not None:
    try:
        df_num = df_bruto.copy()
        col_mes = df_bruto.columns[14]
        df_num[col_mes] = pd.to_numeric(df_num[col_mes], errors='coerce')
        df_filtrada = df_num[df_num[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_bruto.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # ── Colunas Base ──────────────────────────────────────────────────
            df_conferencia['Operacao']   = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[1]]).astype(str)
            df_conferencia['Parte']      = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[62]]).astype(str).str.strip()
            
            # --- TRATAMENTO TIPO DE ENERGIA (ALTERAÇÃO SOLICITADA) ---
            energia_original = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[5]]).astype(str).str.strip()
            
            mapa_energia = {
                'Incentivada-50%': 'Incentivada-I5',
                'Incentivada-100%': 'Incentivada-I1',
                'Incentivada-0%': 'Incentivada-I0',
                'Incentivada-CQ50%': 'Incentivada-CQ5'
            }
            df_conferencia['Tipo Energia'] = energia_original.replace(mapa_energia)
            
            # Regra específica para UFV JACARANDA 1
            mask_jacaranda = df_conferencia['Parte'].str.upper() == 'UFV JACARANDA 1'
            df_conferencia.loc[mask_jacaranda, 'Tipo Energia'] = 'Incentivada-I5'
            # --------------------------------------------------------

            df_conferencia['Contraparte']      = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[6]])
            df_conferencia['CP/LP']            = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[4]]).apply(formatar_cnpj)

            mapa_sub = {'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'}
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[8]]).replace(mapa_sub)

            df_conferencia['Montante MWh'] = pd.to_numeric(
                df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[17]]), errors='coerce'
            ).fillna(0).round(3)

            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[20]]), errors='coerce')
            h_mes = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(6)

            # ── Modulações ────────────────────────────────────────────────────
            df_conferencia['CliqCCEE Paradigma']             = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC']                  = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[63]]).apply(limpar_modulacao)
            df_conferencia['Modulacao Minima']               = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[28]])
            df_conferencia['Modulacao Maxima']               = df_conferencia[col_boleta].map(df_lookup[df_bruto.columns[29]])
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")

            # ── Comprador e Vendedor (do Exportador) ──────────────────────────
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            def resolver(row):
                vend    = row['Vendedor']    if row['Vendedor']    != "-" else ""
                comp    = row['Comprador']   if row['Comprador']   != "-" else ""
                cod_par = row['CliqCCEE Paradigma']
                cod_ant = row['Contrato CliqCCEE mes anterior']
                parte   = str(row['Parte']).upper()

                if 'BISMUT' in parte:
                    return buscar_cliq_ccee(cod_par, cod_ant,
                                            st.session_state['db_bismut'], 'bismut',
                                            vend, comp)

                for tipo, db_key in [('ccear', 'db_ccear'),
                                     ('cbr',   'db_cbr'),
                                     ('matrix','db_matrix')]:
                    resultado = buscar_cliq_ccee(cod_par, cod_ant,
                                                 st.session_state[db_key], tipo,
                                                 vend, comp)
                    if resultado != "Verificar":
                        return resultado

                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver, axis=1)

            # ── Filtros ───────────────────────────────────────────────────────
            st.write("### Filtros")
            f1, f2, f3, f4 = st.columns([2, 2, 2, 1])
            op_f        = f1.selectbox("Operação",          ["Todos"] + sorted(df_conferencia['Operacao'].dropna().unique().tolist()))
            parte_f     = f2.selectbox("Parte",             ["Todos"] + sorted(df_conferencia['Parte'].dropna().unique().tolist()))
            cliq_f      = f3.selectbox("Contrato CliqCCEE", ["Todos"] + sorted(df_conferencia['Contrato CliqCCEE'].dropna().unique().tolist()))
            rem_zero    = f4.toggle("Ocultar Zero",        value=False)
            zerar_intra = f4.toggle("Zerar Intraportfólio", value=False)

            df_final = df_conferencia.copy()
            if op_f    != "Todos": df_final = df_final[df_final['Operacao']           == op_f]
            if parte_f != "Todos": df_final = df_final[df_final['Parte']              == parte_f]
            if cliq_f  != "Todos": df_final = df_final[df_final['Contrato CliqCCEE'] == cliq_f]

            if zerar_intra:
                mask_intra = df_final['Vendedor'].str.strip().str.lower() == df_final['Comprador'].str.strip().str.lower()
                df_final.loc[mask_intra, 'Montante MWh'] = 0.0
                df_final.loc[mask_intra, 'Volume MWm']   = 0.0

            if rem_zero: df_final = df_final[df_final['Volume MWm'] != 0]

            # ── Resumo ────────────────────────────────────────────────────────
            st.write("### Resumo")
            m1, m2, m3 = st.columns(3)
            m1.metric("Contratos", len(df_final))
            m2.metric("Compras",   len(df_final[df_final['Operacao'].str.upper() == 'COMPRA']))
            m3.metric("Vendas",    len(df_final[df_final['Operacao'].str.upper() == 'VENDA']))
            st.markdown("---")

            ordem = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte',
                'Contraparte', 'CP/LP', 'CNPJ Contraparte',
                'Submercado', 'Montante MWh', 'Volume MWm',
                'CliqCCEE Paradigma',
                'Modulacao WBC', 'Modulacao Minima', 'Modulacao Maxima',
                'Contrato CliqCCEE mes anterior', 'Comprador', 'Vendedor',
                'Contrato CliqCCEE',
            ]

            st.dataframe(
                df_final[ordem].sort_values(by=col_boleta),
                hide_index=True, use_container_width=True,
                column_config={
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"),
                    "Volume MWm":   st.column_config.NumberColumn(format="%.6f"),
                }
            )
        else:
            st.warning(f"Sem dados para {mes_nome_sel}/{ano_sel}")
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
