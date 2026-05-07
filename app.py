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
        if df_cliq is None: continue
        if cod in df_cliq.index:
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
        # Só analisamos quem tem volume e não foi identificado automaticamente
        if row['Contrato CliqCCEE'] != "Verificar": continue
        
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

# 6. CARREGAMENTO DOS DADOS (Omitido aqui por brevidade, manter igual ao seu anterior)
# [Mantenha aqui todo o bloco de file_uploader e session_state que já funciona]
# ... (Código de carregamento igual ao anterior) ...

# 7. PROCESSAMENTO DA TABELA
if st.session_state['df_bruto'] is not None:
    try:
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

            # Colunas básicas
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            
            mapa_energia = {'Incentivada-50%': 'Incentivada-I5', 'Incentivada-100%': 'Incentivada-I1', 'Incentivada-0%': 'Incentivada-I0', 'Incentivada-CQ50%': 'Incentivada-CQ5'}
            df_conferencia['Tipo Energia'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[5]]).astype(str).str.strip().replace(mapa_energia)
            
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['CP/LP'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[12]])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[4]]).apply(formatar_cnpj)
            df_conferencia['Submercado'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[8]]).replace({'SE/CO': 'Sudeste', 'N': 'Norte', 'NE': 'Nordeste', 'S': 'Sul'})

            # Volumes
            df_conferencia['Montante MWh'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[17]]), errors='coerce').fillna(0).round(3)
            v_mwh = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[20]]), errors='coerce')
            h_mes_serie = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[df_base.columns[15]]), errors='coerce')
            df_conferencia['Volume MWm'] = (v_mwh / h_mes_serie).fillna(0).round(6)

            # Mapa e Cliq
            df_conferencia['Situacao ERP'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mapa']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")

            # Resoluções CCEE
            def resolver_cliq(row):
                vend, comp = (row['Vendedor'] if row['Vendedor'] != "-" else ""), (row['Comprador'] if row['Comprador'] != "-" else "")
                if 'BISMUT' in str(row['Parte']).upper(): return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear', 'db_ccear'), ('cbr', 'db_cbr'), ('matrix', 'db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)
            
            # --- GERAÇÃO DOS RELATÓRIOS DE MATCH (O QUE HAVIA SUMIDO) ---
            com_match, sem_match, incompleto = gerar_relatorio_match(df_conferencia)

            # Cálculo de Volumes
            df_soma_cliq = df_conferencia[~df_conferencia['Contrato CliqCCEE'].isin(['Verificar', '-', ''])].copy()
            dict_soma_book = df_soma_cliq.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['Volume BOOK'] = df_conferencia['Contrato CliqCCEE'].map(dict_soma_book).fillna(0.0).round(6)

            # (Função buscar_volume_cliq e calculo de Volume CliqCCEE mantidos aqui...)
            # [Mantenha o bloco de Volume CliqCCEE e Validação Volume]
            
            # Lógica de Pagamento
            df_pagos = df_conferencia[df_conferencia['Situacao ERP'].astype(str).str.upper() == 'PAGO'].copy()
            dict_soma_pagos = df_pagos.groupby('Contrato CliqCCEE')['Volume MWm'].sum().to_dict()
            df_conferencia['SITUAÇÃO PGTO'] = df_conferencia.apply(lambda r: "Pago" if r['Contrato CliqCCEE'] in dict_soma_pagos and round(dict_soma_pagos[r['Contrato CliqCCEE']], 6) >= round(r['Volume BOOK'], 6) and r['Volume BOOK'] > 0 else "-", axis=1)

            # ─────────────────────────────────────────────────────────────────
            # EXIBIÇÃO NA TELA
            # ─────────────────────────────────────────────────────────────────
            st.write("---")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Contratos Totais", len(df_conferencia))
            m2.metric("Compras", len(df_conferencia[df_conferencia['Operacao'].str.upper() == 'COMPRA']))
            m3.metric("Vendas", len(df_conferencia[df_conferencia['Operacao'].str.upper() == 'VENDA']))
            m4.metric("Sem Cliq Identificado", len(df_conferencia[df_conferencia['Contrato CliqCCEE'] == "Verificar"]))

            # --- SEÇÃO DE ALERTAS (RELATÓRIO DE MATCH) ---
            if not sem_match.empty:
                with st.expander("⚠️ ATENÇÃO: Boletas sem correspondência na CCEE", expanded=True):
                    st.warning(f"As {len(sem_match)} boletas abaixo possuem Vendedor/Comprador/Submercado que não existem nas bases Cliq carregadas.")
                    st.table(sem_match)

            st.write("### Tabela Principal")
            # [Filtros e st.dataframe final mantidos aqui...]
            # ...
