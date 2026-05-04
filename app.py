import streamlit as st
import pandas as pd
import re

st.set_page_config(layout="wide", page_title="Book de Energia")

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

# SIDEBAR
st.sidebar.title("Configurações")

arquivo_subido = st.sidebar.file_uploader("1. Base Bruta", type=['xlsx', 'xlsm'])
arquivo_anterior = st.sidebar.file_uploader("2. Base Mês Anterior", type=['xlsx'])
arquivo_vendedor = st.sidebar.file_uploader("3. Base Vendedor/Comprador", type=['xlsx'])

st.title("📑 Book de Energia")

# BASE MÊS ANTERIOR
dict_mes_anterior = {}
if arquivo_anterior:
    df_apoio = pd.read_excel(arquivo_anterior)
    dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()

# BASE VENDEDOR / COMPRADOR
dict_vendedor, dict_comprador = {}, {}

if arquivo_vendedor:
    df_vend = pd.read_excel(arquivo_vendedor)

    # 🔥 padroniza chave
    df_vend.iloc[:, 3] = df_vend.iloc[:, 3].astype(str)

    dict_vendedor = pd.Series(df_vend.iloc[:, 2].values, index=df_vend.iloc[:, 3].values).to_dict()
    dict_comprador = pd.Series(df_vend.iloc[:, 1].values, index=df_vend.iloc[:, 3].values).to_dict()

    st.sidebar.success("✅ Base Vendedor/Comprador carregada!")

# PROCESSAMENTO
if arquivo_subido:

    df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')

    # 🔥 padroniza boleta
    df_bruto.iloc[:, 0] = df_bruto.iloc[:, 0].astype(str)

    # 🔥 CRIA BASE ÚNICA (ESSA É A MELHORIA PRINCIPAL)
    df_base = df_bruto.drop_duplicates(subset=[df_bruto.columns[0]]).set_index(df_bruto.columns[0])

    # colunas
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

    # 🚀 AGORA TUDO VEM DIRETO (rápido)
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

    # cálculo otimizado
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

    # 🔥 vendedor/comprador
    df_conferencia['Vendedor'] = df_conferencia['Boleta'].map(dict_vendedor).fillna("-")
    df_conferencia['Comprador'] = df_conferencia['Boleta'].map(dict_comprador).fillna("-")

    # FILTROS
    op = st.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operação'].unique()))
    df_filtrado = df_conferencia.copy()

    if op != "Todos":
        df_filtrado = df_filtrado[df_filtrado['Operação'] == op]

    # EXIBIÇÃO
    st.dataframe(
        df_filtrado,
        use_container_width=True
    )

    st.caption(f"{len(df_filtrado)} linhas exibidas")

else:
    st.info("Aguardando upload das bases.")
