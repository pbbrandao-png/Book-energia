import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURACAO DA PAGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNCOES DE APOIO
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

def carregar_csv_cliq(arquivo):
    if arquivo is None:
        return None
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

def buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, df_cliq):
    if df_cliq is None:
        return "Verificar"

    def checar(codigo):
        codigo = tratar_chave(codigo)
        if not codigo or codigo not in df_cliq.index:
            return False
        row = df_cliq.loc[codigo]
        if isinstance(row, pd.DataFrame):
            row = row.iloc[0]
        situacao = str(row.get('SITUACAO_CONTRATO', '') or '').strip().upper()
        return situacao != 'RASCUNHO'

    if checar(cod_paradigma):
        return tratar_chave(cod_paradigma)
    if checar(cod_mes_anterior):
        return tratar_chave(cod_mes_anterior)
    return "Verificar"

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

st.sidebar.subheader("Período de Referência")
meses_nomes = [
    "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
    "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1
anos = [str(a) for a in range(2024, 2031)]
ano_sel = st.sidebar.selectbox("Ano", anos, index=2)

st.sidebar.markdown("---")
arquivo_subido   = st.sidebar.file_uploader("1. Base do Mês Atual (Excel)", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Mês Anterior_3.xlsx", type=['xlsx'])
arquivo_pessoas  = st.sidebar.file_uploader("3. RelPers_858 (4).xlsx", type=['xlsx'])

st.sidebar.subheader("Bases Cliq CCEE - Matrix/CBR/CCEAR (101457)")
arq_ccear  = st.sidebar.file_uploader("Cliq CCEAR_Q",             type=['xlsx', 'csv'])
arq_cbr    = st.sidebar.file_uploader("Cliq CBR Mercado",         type=['xlsx', 'csv'])
arq_cceal1 = st.sidebar.file_uploader("Cliq CCEAL Firme 101457",  type=['xlsx', 'csv'])

st.sidebar.subheader("Bases Cliq CCEE - Bismut (101475)")
arq_cceal2 = st.sidebar.file_uploader("Cliq CCEAL Firme 101475",  type=['xlsx', 'csv'])

st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel}")

# 4. CARREGAMENTO DOS CSVs CLIQ
dfs_matrix = []
for arq in [arq_ccear, arq_cbr, arq_cceal1]:
    df = carregar_csv_cliq(arq)
    if df is not None:
        dfs_matrix.append(df)
db_matrix = pd.concat(dfs_matrix) if dfs_matrix else None

df_b = carregar_csv_cliq(arq_cceal2)
db_bismut = df_b if df_b is not None else None

# Leitura do mês anterior
dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_apoio = pd.read_excel(arquivo_anterior, header=0, dtype=str)
        primeira_celula = str(df_apoio.iloc[0, 0]).strip().upper()
        if primeira_celula in ["BOLETA", "ID", "CHAVE"]:
            df_apoio = df_apoio.iloc[1:].reset_index(drop=True)
        df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].apply(tratar_chave)
        df_apoio.iloc[:, 1] = df_apoio.iloc[:, 1].apply(tratar_chave)
        dict_mes_anterior = pd.Series(
            df_apoio.iloc[:, 1].values,
            index=df_apoio.iloc[:, 0].values
        ).to_dict()
    except Exception as e:
        st.warning(f"Erro ao carregar mês anterior: {e}")

# Leitura de pessoas (Comprador/Vendedor)
dict_vendedor, dict_comprador = {}, {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        df_pers['chave'] = df_pers.iloc[:, 3].apply(tratar_chave)
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers['chave'].values).to_dict()
        dict_vendedor  = pd.Series(df_pers.iloc[:, 2].values, index=df_pers['chave'].values).to_dict()
    except Exception:
        pass

# 5. PROCESSAMENTO DA BASE PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

        # MAPEAMENTO DE COLUNAS POR ÍNDICE
        col_boleta         = df_bruto.columns[0]
        col_operacao       = df_bruto.columns[1]
        col_tipo_energia   = df_bruto.columns[5]   # COLUNA F
        col_cnpj           = df_bruto.columns[4]
        col_contraparte    = df_bruto.columns[6]
        col_parte_bk       = df_bruto.columns[62]
        col_mes_suprimento = df_bruto.columns[14]
        col_horas_mes      = df_bruto.columns[15]
        col_volume_mwh     = df_bruto.columns[20]
        col_mod_min        = df_bruto.columns[28]
        col_mod_max        = df_bruto.columns[29]
        col_cliq_para      = df_bruto.columns[60]
        col_cliq_alt       = df_bruto.columns[14]
        col_mod_wbc        = df_bruto.columns[63]

        df_bruto[col_mes_suprimento] = pd.to_numeric(df_bruto[col_mes_suprimento], errors='coerce')
        df_filtrada = df_bruto[df_bruto[col_mes_suprimento] == mes_num_sel].copy()

        if df_filtrada.empty:
            st.warning(f"Nenhuma operação encontrada para o mês {mes_num_sel}.")
        else:
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # Preenchimento dos dados com base na Boleta
            df_conferencia['Operação']        = df_conferencia[col_boleta].map(df_lookup[col_operacao]).astype(str)
            df_conferencia['Tipo de Energia'] = df_conferencia[col_boleta].map(df_lookup[col_tipo_energia]).fillna("N/A")
            df_conferencia['Parte']           = df_conferencia[col_boleta].map(df_lookup[col_parte_bk]).astype(str).str.strip()
            df_conferencia['Contraparte']      = df_conferencia[col_boleta].map(df_lookup[col_contraparte])
            df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_lookup[col_cnpj]).apply(formatar_cnpj)

            v_mwh = df_conferencia[col_boleta].map(df_lookup[col_volume_mwh])
            h_mes = df_conferencia[col_boleta].map(df_lookup[col_horas_mes])
            df_conferencia['Volume MWm'] = (v_mwh / h_mes.astype(float)).fillna(0).round(4)

            df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_lookup[col_cliq_para]).apply(tratar_chave)
            df_conferencia['Modulação WBC']      = df_conferencia[col_boleta].map(df_lookup[col_mod_wbc]).apply(limpar_modulacao)
            df_conferencia['Modulação Miníma']   = df_conferencia[col_boleta].map(df_lookup[col_mod_min])
            df_conferencia['Modulação Máxima']   = df_conferencia[col_boleta].map(df_lookup[col_mod_max])

            df_conferencia['Contrato CliqCCEE mes anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")
            df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("N/A")
            df_conferencia['Vendedor']  = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("N/A")

            def resolver_cliq(row):
                parte = str(row['Parte']).strip().upper()
                cod_paradigma    = row['CliqCCEE Paradigma']
                cod_mes_anterior = row['Contrato CliqCCEE mes anterior']
                db = db_bismut if 'BISMUT' in parte else db_matrix
                return buscar_cliq_ccee(cod_paradigma, cod_mes_anterior, db)

            df_conferencia['Contrato CliqCCEE'] = df_conferencia.apply(resolver_cliq, axis=1)

            # Filtros
            lista_op = sorted([str(x) for x in df_conferencia['Operação'].unique() if pd.notna(x)])
            lista_pa = sorted([str(x) for x in df_conferencia['Parte'].unique() if pd.notna(x)])

            st.write("### Filtros Rápidos")
            f1, f2, f3 = st.columns(3)
            with f1:
                op_f = st.selectbox("Operação", ["Todos"] + lista_op)
            with f2:
                parte_f = st.selectbox("Parte", ["Todos"] + lista_pa)
            with f3:
                rem_zero = st.checkbox("Ocultar Zerados", value=False)

            df_final = df_conferencia.copy()
            if op_f != "Todos":
                df_final = df_final[df_final['Operação'] == op_f]
            if parte_f != "Todos":
                df_final = df_final[df_final['Parte'] == parte_f]
            if rem_zero:
                df_final = df_final[df_final['Volume MWm'] != 0]

            # ORDEM DAS COLUNAS (Tipo de Energia como 3ª coluna)
            ordem = [
                col_boleta,
                'Operação',
                'Tipo de Energia',
                'Parte',
                'Contraparte',
                'CNPJ Contraparte',
                'Volume MWm',
                'CliqCCEE Paradigma',
                'Modulação WBC',
                'Modulação Miníma',
                'Modulação Máxima',
                'Contrato CliqCCEE mes anterior',
                'Comprador',
                'Vendedor',
                'Contrato CliqCCEE',
            ]
            
            st.dataframe(df_final[ordem], hide_index=True, use_container_width=True)

    except Exception as e:
        st.error(f"Erro ao processar base principal: {e}")
