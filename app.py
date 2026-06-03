import streamlit as st
import pandas as pd
from io import BytesIO

# =====================================================

# CONFIGURAÇÃO DA PÁGINA

# =====================================================

st.set_page_config(
page_title="Book Energia - Base Conferência",
layout="wide"
)

st.title("📊 Base Conferência")

# =====================================================

# MÊS E ANO

# =====================================================

col1, col2 = st.columns(2)

with col1:
mes = st.selectbox(
"Mês",
[
"Janeiro",
"Fevereiro",
"Março",
"Abril",
"Maio",
"Junho",
"Julho",
"Agosto",
"Setembro",
"Outubro",
"Novembro",
"Dezembro"
]
)

with col2:
ano = st.selectbox(
"Ano",
list(range(2024, 2036)),
index=2
)

# =====================================================

# UPLOAD

# =====================================================

arquivo = st.file_uploader(
"Selecione a RelPers",
type=["xlsx", "xlsm"]
)

# =====================================================

# PROCESSAMENTO

# =====================================================

if arquivo is not None:

```
try:

    # ============================================
    # LEITURA DA PLANILHA
    # ============================================

    df = pd.read_excel(
        arquivo,
        header=8
    )

    # ============================================
    # MAPEAMENTOS
    # ============================================

    mapa_energia = {
        "Incentivada 50%": "Incentivada-I5",
        "Cogeração Qualificada 50%": "Incentivada-CQ5",
        "Incentivada 100%": "Incentivada-I1",
        "Convencional": "Convencional",
        "Incentivada 0%": "Incentivada-I0"
    }

    mapa_submercado = {
        "Sul": "SUL",
        "S": "SUL",
        "SE/CO": "SUDESTE",
        "N": "NORTE",
        "NE": "NORDESTE"
    }

    mapa_modulacao = {
        "F - Flat": "FLAT",
        "C - Carga": "CARGA",
        "DECLARADO": "DECLARADA",
        "G - Geração": "GERAÇÃO"
    }

    # ============================================
    # DATAS
    # ============================================

    df["Suprimento_inicio"] = pd.to_datetime(
        df["Suprimento_inicio"],
        errors="coerce"
    )

    df["Suprimento_termino"] = pd.to_datetime(
        df["Suprimento_termino"],
        errors="coerce"
    )

    dias_periodo = (
        df["Suprimento_termino"]
        - df["Suprimento_inicio"]
    ).dt.days + 1

    horas_periodo = dias_periodo * 24

    # ============================================
    # CP / LP
    # ============================================

    cp_lp = dias_periodo.apply(
        lambda x: "CP" if x <= 31 else "LP"
    )

    # ============================================
    # MWm
    # ============================================

    volume_mwm = (
        df["QuantAtualizada"]
        / horas_periodo
    ).round(4)

    # ============================================
    # BASE FINAL
    # ============================================

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

    base["Contraparte"] = df["Sigla_CCEE_Contraparte"]

    base["CP/LP"] = cp_lp

    base["CNPJ CONTRAPARTE"] = df["Contraparte_CNPJ"]

    base["Submercado"] = (
        df["Submercado"]
        .astype(str)
        .str.strip()
        .map(mapa_submercado)
        .fillna(df["Submercado"])
    )

    base["Volume (MWh)"] = df["QuantAtualizada"]

    base["Volume MWm"] = volume_mwm

    base["CliqCCEE Paradigma"] = df["Codigo_CCEE"]

    base["Modulação WBC"] = (
        df["Tipo_de_modulacao"]
        .astype(str)
        .str.strip()
        .map(mapa_modulacao)
        .fillna(df["Tipo_de_modulacao"])
    )

    base["Modulação Mínima"] = df["FlexLimite_modulacaoMin"]

    base["Modulação Máxima"] = df["FlexLimite_modulacaoMax"]

    # ============================================
    # VISUALIZAÇÃO
    # ============================================

    st.success(
        f"{len(base):,} registros processados."
    )

    st.dataframe(
        base,
        use_container_width=True,
        hide_index=True
    )

    # ============================================
    # DOWNLOAD
    # ============================================

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
        "📥 Download Base Conferência",
        data=output.getvalue(),
        file_name=f"Base_Conferencia_{mes}_{ano}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

except Exception as erro:

    st.error("Erro ao processar a planilha.")

    st.exception(erro)
