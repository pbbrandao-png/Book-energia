import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
page_title="Base Conferência",
layout="wide"
)

st.title("📊 Base Conferência")

# ==========================

# FILTROS

# ==========================

meses = {
"Janeiro": 1,
"Fevereiro": 2,
"Março": 3,
"Abril": 4,
"Maio": 5,
"Junho": 6,
"Julho": 7,
"Agosto": 8,
"Setembro": 9,
"Outubro": 10,
"Novembro": 11,
"Dezembro": 12
}

col1, col2 = st.columns(2)

with col1:
mes = st.selectbox(
"Mês",
list(meses.keys())
)

with col2:
ano = st.selectbox(
"Ano",
list(range(2024, 2036)),
index=2
)

arquivo = st.file_uploader(
"Upload RelPers",
type=["xlsx", "xlsm"]
)

if arquivo:

```
df = pd.read_excel(arquivo)

# ==========================
# MAPAS
# ==========================

mapa_energia = {
    "Incentivada 50%": "Incentivada-I5",
    "Cogeração Qualificada 50%": "Incentivada-CQ5",
    "Incentivada 100%": "Incentivada-I1",
    "Convencional": "Convencional",
    "Incentivada 0%": "Incentivada-I0"
}

mapa_submercado = {
    "N": "NORTE",
    "NE": "NORDESTE",
    "S": "SUL",
    "SE/CO": "SUDESTE"
}

mapa_modulacao = {
    "F": "FLAT",
    "C": "CARGA",
    "DECLARADO": "DECLARADA",
    "G": "GERAÇÃO"
}

# ==========================
# DATAS
# ==========================

df["Suprimento_inicio"] = pd.to_datetime(
    df["Suprimento_inicio"]
)

df["Suprimento_termino"] = pd.to_datetime(
    df["Suprimento_termino"]
)

# ==========================
# CP / LP
# ==========================

dias = (
    df["Suprimento_termino"]
    - df["Suprimento_inicio"]
).dt.days + 1

cp_lp = dias.apply(
    lambda x: "CP" if x <= 31 else "LP"
)

# ==========================
# MWm
# ==========================

horas = dias * 24

volume_mwm = (
    df["QuantAtualizada"] / horas
).round(4)

# ==========================
# BASE FINAL
# ==========================

base = pd.DataFrame()

base["Mês"] = mes
base["Ano"] = ano

base["BOLETA"] = df["Codigo_WBC"]

base["Operação"] = df["Movimentacao"]

base["Tipo de Energia"] = (
    df["Fonte_Contrato"]
    .map(mapa_energia)
    .fillna(df["Fonte_Contrato"])
)

base["Parte"] = df["Parte_razao_social"]

base["Contraparte"] = (
    df["Sigla_CCEE_Contraparte"]
)

base["CP/LP"] = cp_lp

base["CNPJ CONTRAPARTE"] = (
    df["Contraparte_CNPJ"]
)

base["Submercado"] = (
    df["Submercado"]
    .map(mapa_submercado)
    .fillna(df["Submercado"])
)

base["Volume (MWh)"] = (
    df["QuantAtualizada"]
)

base["Volume MWm"] = volume_mwm

base["CliqCCEE Paradigma"] = (
    df["Codigo_CCEE"]
)

base["Modulação WBC"] = (
    df["Tipo_de_modulacao"]
    .astype(str)
    .str.strip()
    .map(mapa_modulacao)
    .fillna(df["Tipo_de_modulacao"])
)

base["Modulação Mínima"] = (
    df["FlexLimite_modulacaoMin"]
)

base["Modulação Máxima"] = (
    df["FlexLimite_modulacaoMax"]
)

st.success(
    f"{len(base):,} registros carregados."
)

st.dataframe(
    base,
    use_container_width=True
)

# ==========================
# DOWNLOAD
# ==========================

output = BytesIO()

with pd.ExcelWriter(
    output,
    engine="openpyxl"
) as writer:

    base.to_excel(
        writer,
        sheet_name="Base Conferência",
        index=False
    )

st.download_button(
    label="📥 Download Base Conferência",
    data=output.getvalue(),
    file_name=f"Base_Conferencia_{mes}_{ano}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
```
