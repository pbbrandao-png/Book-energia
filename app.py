import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Book Energia - Base Conferência", layout="wide")
st.title("📊 Base Conferência e NETS Energéticos")

col1, col2 = st.columns(2)

with col1:
    mes = st.selectbox(
        "Mês",
        ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
         "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
    )

with col2:
    ano = st.selectbox("Ano", list(range(2024, 2036)), index=2)

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx", "xlsm"])

if arquivo is not None:

    try:

        df = pd.read_excel(arquivo, header=8)

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

        df["Suprimento_inicio"] = pd.to_datetime(df["Suprimento_inicio"], errors="coerce")
        df["Suprimento_termino"] = pd.to_datetime(df["Suprimento_termino"], errors="coerce")

        dias_periodo = (df["Suprimento_termino"] - df["Suprimento_inicio"]).dt.days + 1
        horas_periodo = dias_periodo * 24

        cp_lp = dias_periodo.apply(lambda x: "CP" if x <= 31 else "LP")

        volume_mwm = (df["QuantAtualizada"] / horas_periodo).round(4)

        base = pd.DataFrame()

        base["Mês"] = mes
        base["Ano"] = ano
        base["BOLETA"] = df["Codigo_WBC"]
        base["Operação"] = df["Movimentacao"]
        base["Tipo de Energia"] = df["Fonte_Contrato"].map(mapa_energia).fillna(df["Fonte_Contrato"])
        base["Parte"] = df["Parte_razao_social"]
        base["Contraparte"] = df["Sigla_CCEE_Contraparte"]
        base["CP/LP"] = cp_lp
        base["CNPJ CONTRAPARTE"] = df["Contraparte_CNPJ"]
        base["Submercado"] = (
            df["Submercado"].astype(str).str.strip()
            .map(mapa_submercado).fillna(df["Submercado"])
        )
        base["Volume (MWh)"] = df["QuantAtualizada"]
        base["Volume MWm"] = volume_mwm
        base["CliqCCEE Paradigma"] = df["Codigo_CCEE"]
        base["Modulação WBC"] = (
            df["Tipo_de_modulacao"].astype(str).str.strip()
            .map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        )
        base["Modulação Mínima"] = df["FlexLimite_modulacaoMin"]
        base["Modulação Máxima"] = df["FlexLimite_modulacaoMax"]

        st.subheader("Base Conferência")
        st.dataframe(base, use_container_width=True, hide_index=True)

        compras = (
            base[base["Operação"] == "Compra"]
            .groupby(
                ["Parte","Contraparte","Submercado","Tipo de Energia"],
                as_index=False
            )["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)": "Compra (MWh)"})
        )

        vendas = (
            base[base["Operação"] == "Venda"]
            .groupby(
                ["Parte","Contraparte","Submercado","Tipo de Energia"],
                as_index=False
            )["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)": "Venda (MWh)"})
        )

        nets = compras.merge(
            vendas,
            on=["Parte","Contraparte","Submercado","Tipo de Energia"],
            how="inner"
        )

        nets["NET (MWh)"] = nets["Compra (MWh)"] - nets["Venda (MWh)"]

        def registrante(row):
            if row["NET (MWh)"] > 0:
                return row["Contraparte"]
            elif row["NET (MWh)"] < 0:
                return row["Parte"]
            return "ZERADO"

        nets["Registrante"] = nets.apply(registrante, axis=1)

        st.subheader("NETS ENERGÉTICOS")

        parte_filtro = st.selectbox(
            "Filtrar Parte",
            ["Todos"] + sorted(nets["Parte"].dropna().unique().tolist())
        )

        if parte_filtro != "Todos":
            nets = nets[nets["Parte"] == parte_filtro]

        contraparte_filtro = st.selectbox(
            "Filtrar Contraparte",
            ["Todos"] + sorted(nets["Contraparte"].dropna().unique().tolist())
        )

        if contraparte_filtro != "Todos":
            nets = nets[nets["Contraparte"] == contraparte_filtro]

        st.dataframe(nets, use_container_width=True, hide_index=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            base.to_excel(writer, sheet_name="Base Conferência", index=False)
            nets.to_excel(writer, sheet_name="NETS ENERGÉTICOS", index=False)

        st.download_button(
            "📥 Download Excel",
            data=output.getvalue(),
            file_name=f"Book_Energia_{mes}_{ano}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)
