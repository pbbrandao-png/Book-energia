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
# MAPEAMENTO E REGRAS DE BUSCA CLIQ
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
        try:
            row = df_cliq.loc[codigo]
            if isinstance(row, pd.DataFrame): row = row.iloc[0]
            if str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper() == 'RASCUNHO': return False
            if col_vend and col_vend in df_cliq.columns:
                if limpar_str(nome_vendedor) and limpar_str(row.get(col_vend, '')) != limpar_str(nome_vendedor): return False
            if col_comp and col_comp in df_cliq.columns:
                if limpar_str(nome_comprador) and limpar_str(row.get(col_comp, '')) != limpar_str(nome_comprador): return False
            return True
        except: return False

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

# 4. INTERFACE LATERAL (Omitido carregamento manual aqui por brevidade, manter igual ao anterior)
# ... [Omitido para focar na lógica de processamento corrigida] ...

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

            # [Colunas Básicas Mapeadas como nas versões anteriores...]
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[6]])
            df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_vendedor']).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_comprador']).fillna("-")
            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[60]]).apply(tratar_chave)
            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(st.session_state['dict_mes_anterior']).fillna("-")
            # [Adicione as demais colunas de volume, cnpj, etc aqui...]

            # Resolver o Código Cliq
            def resolver_cliq(row):
                vend, comp = (row['Vendedor'] if row['Vendedor'] != "-" else ""), (row['Comprador'] if row['Comprador'] != "-" else "")
                if 'BISMUT' in str(row['Parte']).upper(): 
                    return buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state['db_bismut'], 'bismut', vend, comp)
                for t, k in [('ccear','db_ccear'), ('cbr','db_cbr'), ('matrix','db_matrix')]:
                    res = buscar_cliq_ccee(row['CliqCCEE Paradigma'], row['Contrato CliqCCEE mes anterior'], st.session_state[k], t, vend, comp)
                    if res != "Verificar": return res
                return "Verificar"
            
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)

            # --- BUSCA BLINDADA COM REGRA CCEAR Q ---
            def buscar_info_cliq(row, campo):
                cod = row['Contrato CliqCCEE']
                if cod in ['Verificar', '-', '']: return "-"
                
                for db_key in ['db_matrix', 'db_bismut', 'db_ccear', 'db_cbr']:
                    df_cliq = st.session_state.get(db_key)
                    if df_cliq is not None and cod in df_cliq.index:
                        # Regra específica para Status Montante na CCEAR_Q
                        if campo == 'STATUS_MONTANTE' and db_key == 'db_ccear':
                            return "AJUSTE VALIDADO"
                        
                        try:
                            val = df_cliq.loc[cod, campo]
                            if isinstance(val, pd.Series): val = val.iloc[0]
                            return str(val).strip() if not pd.isna(val) else "-"
                        except:
                            # Se der erro na CCEAR para o campo STATUS_MONTANTE, aplica o padrão
                            if campo == 'STATUS_MONTANTE' and db_key == 'db_ccear': return "AJUSTE VALIDADO"
                            continue
                return "-"
            
            df_conferencia['Status do Contrato'] = df_conferencia.apply(lambda r: buscar_info_cliq(r, 'SITUACAO_CONTRATO'), axis=1)
            df_conferencia['Status Montante'] = df_conferencia.apply(lambda r: buscar_info_cliq(r, 'STATUS_MONTANTE'), axis=1)

            # [Resto da lógica de volumes, filtros e exibição igual às versões anteriores...]
            # ...

            st.dataframe(df_conferencia, use_container_width=True, hide_index=True)
    except Exception as e: st.error(f"Erro no processamento: {e}")
