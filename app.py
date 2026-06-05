# APP_BOOK_ENERGIA_V6
# Removidas colunas Mês e Ano
# CNPJ formatado
# MWh = 3 casas
# MWm = 6 casas

import streamlit as st
import pandas as pd
from io import BytesIO

def formatar_cnpj(valor):
    if pd.isna(valor):
        return ""

    cnpj = "".join(filter(str.isdigit, str(valor)))
    cnpj = cnpj.zfill(14)

    return (
        f"{cnpj[:2]}."
        f"{cnpj[2:5]}."
        f"{cnpj[5:8]}/"
        f"{cnpj[8:12]}-"
        f"{cnpj[12:]}"
    )

st.set_page_config(page_title="Book Energia", layout="wide")

pagina = st.sidebar.radio(
    "Menu",
    ["Base Conferência", "Encontro Energético"]
)

st.title("📊 Book Energia")

arquivo = st.file_uploader(
    "Selecione a RelPers",
    type=["xlsx", "xlsm"]
)

arquivo_mes_anterior = st.file_uploader(
    "Selecione a planilha Mês Anterior",
    type=["xlsx"]
)

if arquivo is not None:

    try:

        df = pd.read_excel(arquivo, header=8)

        horas_mes = {
            1: 744,
            2: 672,
            3: 744,
            4: 720,
            5: 744,
            6: 720,
            7: 744,
            8: 744,
            9: 720,
            10: 744,
            11: 720,
            12: 744
        }

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)

            mapa_mes_anterior = dict(
                zip(
                    df_mes_anterior["BOLETA"],
                    df_mes_anterior["Codigo_CCEE"]
                )
            )
        else:
            mapa_mes_anterior = {}

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

        cp_lp = dias_periodo.apply(lambda x: "CP" if x <= 31 else "LP")

        horas_por_linha = df["Mes"].map(horas_mes)

        volume_mwm = (
            df["QuantAtualizada"] / horas_por_linha
        ).round(6)

        base = pd.DataFrame()

        base["BOLETA"] = df["Codigo_WBC"]
        base["Operação"] = df["Movimentacao"]
        base["Tipo de Energia"] = df["Fonte_Contrato"].map(mapa_energia).fillna(df["Fonte_Contrato"])
        base["Parte"] = df["Parte_razao_social"]
        base["Contraparte"] = df["Sigla_CCEE_Contraparte"]
        base["CP/LP"] = cp_lp
        base["CNPJ CONTRAPARTE"] = df["Contraparte_CNPJ"].apply(formatar_cnpj)
        base["Submercado"] = df["Submercado"].astype(str).str.strip().map(mapa_submercado).fillna(df["Submercado"])
        base["Volume (MWh)"] = df["QuantAtualizada"].round(3)
        base["Volume MWm"] = volume_mwm.round(6)
        base["CliqCCEE Paradigma"] = df["Codigo_CCEE"]
        base["Modulação WBC"] = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["Modulação Mínima"] = df["FlexLimite_modulacaoMin"].fillna("-")
        base["Modulação Máxima"] = df["FlexLimite_modulacaoMax"].fillna("-")

        base["Contrato CliqCCEE mês anterior"] = (
            base["BOLETA"]
            .map(mapa_mes_anterior)
            .fillna("-")
        )

        base["Vendedor"] = df["Sigla_CCEE_vendedor"]
        base["Comprador"] = df["Sigla_CCEE_comprador"]

        compras_net = (
            base[base["Operação"] == "Compra"]
            .groupby(["Parte","Contraparte","Submercado","Tipo de Energia"], as_index=False)["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)":"Compra (MWh)"})
        )

        vendas_net = (
            base[base["Operação"] == "Venda"]
            .groupby(["Parte","Contraparte","Submercado","Tipo de Energia"], as_index=False)["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)":"Venda (MWh)"})
        )

        nets = compras_net.merge(
            vendas_net,
            on=["Parte","Contraparte","Submercado","Tipo de Energia"],
            how="inner"
        )

        if pagina == "Base Conferência":

            st.subheader("Base Conferência")

            base_exibicao = base.copy()
            base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            base_exibicao["Volume MWm"] = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}")

            st.dataframe(base_exibicao, use_container_width=True, hide_index=True)

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                base.to_excel(writer, sheet_name="Base Conferência", index=False)

            st.download_button(
                "📥 Download Base Conferência",
                data=output.getvalue(),
                file_name="Base_Conferencia.xlsx"
            )

        else:

            st.subheader("🤝 Encontro Energético")

            parte = st.selectbox("Parte", sorted(nets["Parte"].dropna().unique()))
            df_parte = nets[nets["Parte"] == parte]

            contraparte = st.selectbox("Contraparte", sorted(df_parte["Contraparte"].dropna().unique()))
            df_contraparte = df_parte[df_parte["Contraparte"] == contraparte]

            submercado = st.selectbox("Submercado", sorted(df_contraparte["Submercado"].dropna().unique()))
            df_sub = df_contraparte[df_contraparte["Submercado"] == submercado]

            tipo_energia = st.selectbox("Tipo de Energia", sorted(df_sub["Tipo de Energia"].dropna().unique()))

            encontro = base[
                (base["Parte"] == parte) &
                (base["Contraparte"] == contraparte) &
                (base["Submercado"] == submercado) &
                (base["Tipo de Energia"] == tipo_energia)
            ]

            compras_calc = encontro[encontro["Operação"] == "Compra"]
            vendas_calc = encontro[encontro["Operação"] == "Venda"]

            compras = compras_calc.copy()
            vendas = vendas_calc.copy()

            compras["Volume (MWh)"] = compras["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            compras["Volume MWm"] = compras["Volume MWm"].map(lambda x: f"{x:.6f}")

            vendas["Volume (MWh)"] = vendas["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            vendas["Volume MWm"] = vendas["Volume MWm"].map(lambda x: f"{x:.6f}")

            st.markdown("## COMPRAS")
            st.dataframe(compras[["BOLETA","Volume (MWh)","Volume MWm"]], hide_index=True, use_container_width=True)

            st.markdown("## VENDAS")
            st.dataframe(vendas[["BOLETA","Volume (MWh)","Volume MWm"]], hide_index=True, use_container_width=True)

            total_compra = compras_calc["Volume (MWh)"].sum()
            total_venda = vendas_calc["Volume (MWh)"].sum()
            saldo = total_compra - total_venda

            total_compra_mwm = compras_calc["Volume MWm"].sum()
            total_venda_mwm = vendas_calc["Volume MWm"].sum()

            mes_referencia = int(df["Mes"].dropna().iloc[0])
            saldo_mwm = saldo / horas_mes.get(mes_referencia, 744)

            ajuste = contraparte if saldo > 0 else parte if saldo < 0 else "ZERADO"

            resumo = pd.DataFrame({
                "Tipo":["Compras","Vendas","Saldo"],
                "MWh":[f"{total_compra:.3f}", f"{total_venda:.3f}", f"{saldo:.3f}"],
                "MWm":[f"{total_compra_mwm:.6f}", f"{total_venda_mwm:.6f}", f"{saldo_mwm:.6f}"]
            })

            st.markdown("## RESUMO")
            st.dataframe(resumo, hide_index=True, use_container_width=True)

            c1, c2 = st.columns(2)

            with c1:
                st.metric("Quem Ajusta", ajuste)

            with c2:
                st.metric("Volume a Ajustar (MWm)", f"{abs(saldo_mwm):.6f}")

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)
