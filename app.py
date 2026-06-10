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
    ["Base Conferência", "Encontro Energético", "Arquivos CCEE", "Declaradas"]
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

zip_matrix = st.file_uploader(
    "Selecione o ZIP Matrix",
    type=["zip"]
)

zip_bismut = st.file_uploader(
    "Selecione o ZIP Bismut",
    type=["zip"]
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

        elif pagina == "Encontro Energético":

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

        elif pagina == "Arquivos CCEE":
            st.subheader("📁 Arquivos CCEE")
            st.info("Em construção.")

        elif pagina == "Declaradas":

            st.subheader("📋 Declaradas")

            arquivo_xml = st.file_uploader(
                "Selecione o XML da CCEE (Declaração de Modulação)",
                type=["xml"],
                key="xml_declaradas"
            )

            arquivo_med = st.file_uploader(
                "Selecione o arquivo de Medições (xlsx)",
                type=["xlsx"],
                key="med_declaradas"
            )

            if arquivo_xml is not None and arquivo_med is not None:

                try:
                    import xml.etree.ElementTree as ET

                    # ── Lê XML ──────────────────────────────────────────────
                    tree = ET.parse(arquivo_xml)
                    root_xml = tree.getroot()

                    contrato_node = root_xml.find("Contrato")
                    numero_contrato = contrato_node.attrib.get("numeroContrato", "")

                    mm_node = contrato_node.find("MontanteMedio")
                    firme_node = mm_node.find("MontanteMedioContratoCCEALFirme")
                    montante_medio_mwm = float(firme_node.attrib["montanteMedio"])

                    vigencia_inicio = mm_node.attrib.get("vigenciaDeInicio", "")
                    vigencia_fim    = mm_node.attrib.get("vigenciaDeFim", "")

                    # Extrai mês/ano e horas totais do período
                    mesano_node = mm_node.find("MesAno")
                    mes_ano_str = mesano_node.attrib["mesAno"]           # "03/2026"
                    mes_ref = int(mes_ano_str.split("/")[0])
                    horas_periodo = horas_mes.get(mes_ref, 744)

                    # MWh total contratado = MWm médio × horas do período
                    montante_total_mwh = montante_medio_mwm * horas_periodo

                    # Lê todas as horas declaradas do XML → lista de dicts
                    rows_xml = []
                    for dia_node in mesano_node.findall("Dia"):
                        dia = int(dia_node.attrib["dia"])
                        for hora_node in dia_node.findall("Hora"):
                            hora = int(hora_node.attrib["hora"])          # 0–23
                            mwm  = float(hora_node.attrib["montanteHorario"])
                            rows_xml.append({
                                "Dia": dia,
                                "Hora_XML": hora,     # 0-based (padrão CCEE)
                                "Declarado_MWm": mwm,
                                "Declarado_MWh": mwm  # 1 hora → MWh = MWm × 1
                            })

                    df_xml = pd.DataFrame(rows_xml)

                    # ── Lê Medições ─────────────────────────────────────────
                    df_med = pd.read_excel(arquivo_med, header=5)
                    df_med.columns = [
                        "Agente", "Ponto", "Data", "Hora",
                        "Ativa_kWh", "Qualidade", "Origem"
                    ]

                    # Hora nas medições é 1-based; converte para 0-based p/ merge
                    df_med["Hora_XML"] = df_med["Hora"] - 1

                    # Data pode vir como datetime ou serial Excel
                    if pd.api.types.is_datetime64_any_dtype(df_med["Data"]):
                        df_med["Dia"] = df_med["Data"].dt.day
                    else:
                        # serial Excel → datetime
                        df_med["Dia"] = pd.to_datetime(
                            df_med["Data"], unit="D", origin="1899-12-30"
                        ).dt.day

                    # Agrega por dia+hora (caso haja múltiplos pontos de medição)
                    agg_med = (
                        df_med
                        .groupby(["Dia", "Hora_XML"], as_index=False)["Ativa_kWh"]
                        .sum()
                    )
                    agg_med["Medido_MWh"] = agg_med["Ativa_kWh"] / 1000
                    agg_med["Medido_MWm"] = agg_med["Medido_MWh"]   # 1h → MWm = MWh

                    # ── Merge declarado × medido ─────────────────────────────
                    df_dec = df_xml.merge(
                        agg_med[["Dia", "Hora_XML", "Medido_MWh", "Medido_MWm"]],
                        on=["Dia", "Hora_XML"],
                        how="left"
                    )

                    df_dec["Hora_Exibicao"] = df_dec["Hora_XML"]   # mantém 0-based (CCEE)
                    df_dec["Diferença_MWh"] = (
                        df_dec["Declarado_MWh"] - df_dec["Medido_MWh"]
                    ).round(6)

                    # ── Totais para validação CK vs CP ───────────────────────
                    total_declarado_mwh = df_dec["Declarado_MWh"].sum()   # coluna CP
                    total_declarado_mwm = montante_medio_mwm              # coluna CK
                    total_medido_mwh    = df_dec["Medido_MWh"].sum()

                    # CP (MWh) ÷ horas = deve bater com CK (MWm)
                    cp_convertido_mwm = total_declarado_mwh / horas_periodo

                    ok_ck_cp = abs(cp_convertido_mwm - total_declarado_mwm) < 0.0001

                    # ── Cabeçalho informativo ────────────────────────────────
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Contrato", numero_contrato)
                    col2.metric("Período", f"{vigencia_inicio} → {vigencia_fim}")
                    col3.metric("Mês de referência", mes_ano_str)

                    st.markdown("---")

                    col4, col5, col6 = st.columns(3)
                    col4.metric(
                        "Total Declarado — CK (MWm)",
                        f"{total_declarado_mwm:.6f}"
                    )
                    col5.metric(
                        "Total Declarado — CP (MWh)",
                        f"{total_declarado_mwh:.3f}"
                    )
                    col6.metric(
                        "CP ÷ horas = MWm",
                        f"{cp_convertido_mwm:.6f}",
                        delta="✅ Bate com CK" if ok_ck_cp else "⚠️ Diverge de CK",
                        delta_color="normal" if ok_ck_cp else "inverse"
                    )

                    st.markdown("---")
                    col7, col8 = st.columns(2)
                    col7.metric("Total Medido (MWh)", f"{total_medido_mwh:.3f}")
                    col8.metric(
                        "Diferença Total (MWh)",
                        f"{(total_declarado_mwh - total_medido_mwh):.3f}"
                    )

                    # ── Tabela horária ───────────────────────────────────────
                    st.markdown("### Detalhamento Horário")

                    tabela = df_dec[[
                        "Dia", "Hora_Exibicao",
                        "Declarado_MWm", "Declarado_MWh",
                        "Medido_MWh", "Diferença_MWh"
                    ]].copy()

                    tabela.columns = [
                        "Dia", "Hora (0-based)",
                        "Declarado MWm (CK)", "Declarado MWh (CP)",
                        "Medido MWh", "Diferença MWh"
                    ]

                    st.dataframe(
                        tabela.style.format({
                            "Declarado MWm (CK)": "{:.6f}",
                            "Declarado MWh (CP)": "{:.6f}",
                            "Medido MWh":         "{:.3f}",
                            "Diferença MWh":      "{:.6f}",
                        }),
                        use_container_width=True,
                        hide_index=True
                    )

                    # ── Download ─────────────────────────────────────────────
                    output_dec = BytesIO()
                    with pd.ExcelWriter(output_dec, engine="openpyxl") as writer:
                        tabela.to_excel(
                            writer, sheet_name="Declaradas", index=False
                        )

                    st.download_button(
                        "📥 Download Declaradas",
                        data=output_dec.getvalue(),
                        file_name=f"Declaradas_{numero_contrato}_{mes_ano_str.replace('/','_')}.xlsx"
                    )

                except Exception as erro_dec:
                    st.error("Erro ao processar Declaradas")
                    st.exception(erro_dec)

            else:
                st.info("Faça upload do XML da CCEE e do arquivo de Medições para continuar.")

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)


# =========================
# V15 - MELHORIAS SUGERIDAS
# =========================
# Imports adicionais:
# import win32com.client as win32
#
# Adicionar após cálculo do NET:
#
# nets["NET (MWh)"] = nets["Compra (MWh)"] - nets["Venda (MWh)"]
# nets["NET (MWm)"] = nets["NET (MWh)"] / horas_mes.get(mes_referencia,744)
#
# Criar relatório consolidado e botão Outlook.
#
# email_destino = st.text_input("Para")
#
# if st.button("📧 Gerar E-mail Outlook"):
#     outlook = win32.Dispatch("Outlook.Application")
#     mail = outlook.CreateItem(0)
#     mail.To = email_destino
#     mail.Subject = f"Encontro Energético {mes_referencia}/2026 - {contraparte}"
#     mail.HTMLBody = html_email
#     mail.Display()
