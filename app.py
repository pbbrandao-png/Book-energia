import streamlit as st
import pandas as pd
import re

# CONFIG
st.set_page_config(layout="wide", page_title="Book de Energia")

# FUNÇÕES
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "":
        return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj))
    apenas_numeros = apenas_numeros.zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def limpar_modulacao(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper()
    if "FLAT" in t: return "Flat"
    if "CARGA" in t: return "Carga"
    if "DECLARADO" in t or "INFORMADO" in t: return "Declarado"
    if "GERA" in t: return "Geração"
    return texto

def limpar_chave(valor):
    return str(valor).strip() if pd.notna(valor) else ""

# SIDEBAR
st.sidebar.title("Configurações")

arquivo_subido = st.sidebar.file_uploader("1. Base Bruta", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior", type=['xlsx'])
arquivo_vendedor = st.sidebar.file_uploader("3. Base Vendedor/Comprador", type=['xlsx'])

st.title("📑 Book de Energia")

# BASE ANTERIOR
dict_mes_anterior = {}
if arquivo_anterior:
    df_apoio = pd.read_excel(arquivo_anterior)
    df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].apply(limpar_chave)
    dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()

# BASE VENDEDOR / COMPRADOR
dict_vendedor, dict_comprador = {}, {}

if arquivo_vendedor:
    df_vend = pd.read_excel(arquivo_vendedor)

    # D = Boleta
    df_vend.iloc[:, 3] = df_vend.iloc[:, 3].apply(limpar_chave)

    dict_vendedor = pd.Series(df_vend.iloc[:, 2].values, index=df_vend.iloc[:, 3].values).to_dict()
    dict_comprador = pd.Series(df_vend.iloc[:, 1].values, index=df_vend.iloc[:, 3].values).to_dict()

    st.sidebar.success("✅ Base Vendedor/Comprador carregada!")

# PROCESSAMENTO PRINCIPAL
if arquivo_subido:

    df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

    # 🔥 CORREÇÃO PRINCIPAL (sem erro de dtype)
    df_bruto.iloc[:, 0] = df_bruto.iloc[:, 0].apply(limpar_chave)

    # 🚀 OTIMIZAÇÃO (base única)
    df_base = df_bruto.drop_duplicates(subset=[df_bruto.columns[0]]).set_index(df_bruto.columns[0])

    # COLUNAS
    col_operacao = df_bruto.columns[1]
    col_cnpj = df_bruto.columns[4]
    col_energia = df_bruto.columns[5]
    col_contraparte = df_bruto.columns[6]
    col_horas_mes = df_bruto.columns[15]
    col_volume_mwh = df_bruto.columns[20]
    col_mod_min = df_bruto.columns[28]
    col_mod_max = df_bruto.columns[29]
    col_cliq_para = df_bruto.columns[60]
    col_parte = df_bruto.columns[62]
    col_mod_wbc = df_bruto.columns[63]

    df_conferencia = pd.DataFrame(df_base.index, columns=['Boleta'])

    # DADOS
    df_conferencia['Operação'] = df_base[col_operacao].values

    trad_en = {
        "Incentivada-50%": "Incentivada-I5",
        "Incentivada-CQ50%": "Incentivada-CQ5",
        "Incentivada-100%": "Incentivada-I1",
        "Incentivada-0%": "Incentivada-I0",
        "Convencional": "Convencional"
    }

    df_conferencia['Tipo de Energia'] = df_base[col_energia].replace(trad_en).values
    df_conferencia['Parte'] = df_base[col_parte].astype(str).values
    df_conferencia['Contraparte'] = df_base[col_contraparte].values
    df_conferencia['CNPJ Contraparte'] = df_base[col_cnpj].apply(formatar_cnpj).values

    df_conferencia['Volume MWm'] = (
        (df_base[col_volume_mwh] / df_base[col_horas_mes])
        .fillna(0)
        .round(4)
        .values
    )

    df_conferencia['CliqCCEE Paradigma'] = df_base[col_cliq_para].values
    df_conferencia['Modulação WBC'] = df_base[col_mod_wbc].apply(limpar_modulacao).values
    df_conferencia['Modulação Mínima'] = df_base[col_mod_min].values
    df_conferencia['Modulação Máxima'] = df_base[col_mod_max].values

    df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta'].map(dict_mes_anterior).fillna("-")

    # 🔥 NOVO (Vendedor / Comprador)
    df_conferencia['Vendedor'] = df_conferencia['Boleta'].map(dict_vendedor).fillna("-")
    df_conferencia['Comprador'] = df_conferencia['Boleta'].map(dict_comprador).fillna("-")

    # FILTROS
    st.write("### Filtros")
    col1, col2 = st.columns(2)

    with col1:
        op = st.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operação'].astype(str).unique()))

    with col2:
        parte = st.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].astype(str).unique()))

    df_filtrado = df_conferencia.copy()

    if op != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Operação'] == op]

    if parte != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Parte'] == parte]

    # EXIBIÇÃO
    st.dataframe(
        df_filtrado,
        use_container_width=True
    )

    st.caption(f"Mostrando {len(df_filtrado)} registros")

else:
    st.info("Aguardando upload das bases.")
