# APP_BOOK_ENERGIA_V16
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)

import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

# Boletas que devem buscar no CSV ccear_q em vez do cceal_firme
BOLETAS_ACR = {
    122387, 122389, 122391, 122393, 122395, 122397, 122399, 122401,
    144795, 144797, 144799, 148084, 148088, 148090, 148092, 148518,
}


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


def ler_csv_ccee(bytes_csv):
    """Lê bytes de um CSV CCEE (sep=TAB, encoding=latin1, pula linha sep=;)."""
    df = pd.read_csv(BytesIO(bytes_csv), sep='\t', encoding='latin1', skiprows=1, dtype=str)
    df.columns = df.columns.str.strip()
    for col in ['CODIGO_CONTRATO', 'SITUACAO_CONTRATO',
                'SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA']:
        if col in df.columns:
            df[col] = df[col].str.strip()
    df['_CHAVE'] = (
        df['SIGLA_PERFIL_VENDEDOR'].fillna('')
        + df['SIGLA_PERFIL_COMPRADOR'].fillna('')
        + df['SUBMERCADO_ENTREGA'].fillna('')
    )
    return df


def extrair_csvs_zip(zip_file):
    """
    Extrai do ZIP os DataFrames CCEE relevantes.
    Retorna dict com chaves: 'cceal', 'cbr', 'ccear_q'
    (cada um é um DataFrame ou None se não encontrado no ZIP).
    """
    result = {'cceal': None, 'cbr': None, 'ccear_q': None}
    if zip_file is None:
        return result
    try:
        with zipfile.ZipFile(zip_file) as zf:
            for nome in zf.namelist():
                nome_lower = nome.lower()
                # ignora diretórios
                if nome_lower.endswith('/'):
                    continue
                if not nome_lower.endswith('.csv'):
                    continue
                # arquivos _parcela não têm as colunas necessárias
                if 'parcela' in nome_lower:
                    continue
                dados = zf.read(nome)
                if 'ccear_q' in nome_lower:
                    result['ccear_q'] = ler_csv_ccee(dados)
                elif 'cbr_mercado_proprio' in nome_lower or 'cbr_mercado' in nome_lower:
                    result['cbr'] = ler_csv_ccee(dados)
                elif 'cceal_firme' in nome_lower or 'cceal' in nome_lower:
                    result['cceal'] = ler_csv_ccee(dados)
    except Exception as e:
        st.warning(f"Erro ao ler ZIP: {e}")
    return result


def combinar_dfs(lista):
    """Concatena DataFrames não-nulos da lista."""
    validos = [df for df in lista if df is not None and not df.empty]
    if not validos:
        return pd.DataFrame()
    return pd.concat(validos, ignore_index=True)


def buscar_contrato_cliqccee(codigo_busca, chave_esperada, df_ccee):
    """
    Procura codigo_busca na coluna CODIGO_CONTRATO do df_ccee.
    Retorna o código se bater, 'Verificar' se a chave não conferir, '-' se não achar.
    """
    if df_ccee.empty or pd.isna(codigo_busca) or str(codigo_busca).strip() in ('', '-', 'None'):
        return '-'
    try:
        codigo_busca = str(codigo_busca).strip()
        encontrado = df_ccee[df_ccee['CODIGO_CONTRATO'] == codigo_busca]
        if encontrado.empty:
            return '-'
        row = encontrado.iloc[0]
        situacao = str(row.get('SITUACAO_CONTRATO', '')).strip().lower()
        if situacao == 'rascunho':
            return '-'
        if row['_CHAVE'] == chave_esperada:
            return codigo_busca
        return 'Verificar'
    except Exception:
        return '-'


def resolver_contrato_cliqccee(boleta, codigo_mes_anterior, codigo_paradigma,
                                chave, df_matrix, df_bismut, df_acr, is_bismut):
    """
    Roteamento por tipo de boleta/parte:
      - Boleta ACR    → df_acr  (ccear_q do ZIP Matrix)
      - Bismut        → df_bismut (cceal do ZIP Bismut)
      - Demais        → df_matrix (cceal + cbr do ZIP Matrix)
    Tenta mês anterior primeiro; se 'Verificar', faz fallback pelo paradigma.
    """
    try:
        boleta_int = int(float(str(boleta).strip()))
    except (ValueError, TypeError):
        boleta_int = -1

    if boleta_int in BOLETAS_ACR:
        df = df_acr
    elif is_bismut:
        df = df_bismut
    else:
        df = df_matrix

    resultado = buscar_contrato_cliqccee(codigo_mes_anterior, chave, df)
    if resultado == 'Verificar':
        resultado = buscar_contrato_cliqccee(codigo_paradigma, chave, df)
    return resultado


# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Book Energia", layout="wide")

pagina = st.sidebar.radio(
    "Menu",
    ["Base Conferência", "Encontro Energético", "Arquivos CCEE"]
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
            1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720,
            7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744
        }

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
            mapa_mes_anterior = dict(
                zip(df_mes_anterior["BOLETA"], df_mes_anterior["Codigo_CCEE"])
            )
        else:
            mapa_mes_anterior = {}

        # Extrai CSVs dos ZIPs
        csvs_matrix = extrair_csvs_zip(zip_matrix)
        csvs_bismut = extrair_csvs_zip(zip_bismut)

        # DataFrame Matrix: cceal_firme + cbr_mercado_proprio do ZIP Matrix
        df_ccee_matrix = combinar_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        # DataFrame Bismut: cceal_firme do ZIP Bismut
        df_ccee_bismut = combinar_dfs([csvs_bismut['cceal']])
        # DataFrame ACR: ccear_q do ZIP Matrix
        df_ccee_acr = combinar_dfs([csvs_matrix['ccear_q']])

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
        volume_mwm = (df["QuantAtualizada"] / horas_por_linha).round(6)

        base = pd.DataFrame()

        base["BOLETA"]                         = df["Codigo_WBC"]
        base["Operação"]                       = df["Movimentacao"]
        base["Tipo de Energia"]                = df["Fonte_Contrato"].map(mapa_energia).fillna(df["Fonte_Contrato"])
        base["Parte"]                          = df["Parte_razao_social"]
        base["Contraparte"]                    = df["Sigla_CCEE_Contraparte"]
        base["CP/LP"]                          = cp_lp
        base["CNPJ CONTRAPARTE"]               = df["Contraparte_CNPJ"].apply(formatar_cnpj)
        base["Submercado"]                     = df["Submercado"].astype(str).str.strip().map(mapa_submercado).fillna(df["Submercado"])
        base["Volume (MWh)"]                   = df["QuantAtualizada"].round(3)
        base["Volume MWm"]                     = volume_mwm.round(6)
        base["CliqCCEE Paradigma"]             = df["Codigo_CCEE"]
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["Modulação Mínima"]               = df["FlexLimite_modulacaoMin"].fillna("-")
        base["Modulação Máxima"]               = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-")
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"]
        base["Comprador"]                      = df["Sigla_CCEE_comprador"]

        # ── Coluna "Contrato CliqCCEE" ─────────────────────────────────────────
        BISMUT_NOME = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."

        csvs_disponiveis = any([
            not df_ccee_matrix.empty,
            not df_ccee_bismut.empty,
            not df_ccee_acr.empty,
        ])

        if csvs_disponiveis:
            def calcular_contrato_cliqccee(row):
                is_bismut = str(row["Parte"]).strip().upper() == BISMUT_NOME.upper()
                chave = (
                    str(row["Vendedor"]).strip()
                    + str(row["Comprador"]).strip()
                    + str(row["Submercado"]).strip()
                )
                return resolver_contrato_cliqccee(
                    boleta              = row["BOLETA"],
                    codigo_mes_anterior = row["Contrato CliqCCEE mês anterior"],
                    codigo_paradigma    = row["CliqCCEE Paradigma"],
                    chave               = chave,
                    df_matrix           = df_ccee_matrix,
                    df_bismut           = df_ccee_bismut,
                    df_acr              = df_ccee_acr,
                    is_bismut           = is_bismut,
                )
            base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee, axis=1)
        else:
            base["Contrato CliqCCEE"] = "-"
            if pagina == "Base Conferência":
                st.info("ℹ️ Faça upload dos ZIPs para preencher a coluna 'Contrato CliqCCEE'.")
        # ───────────────────────────────────────────────────────────────────────

        compras_net = (
            base[base["Operação"] == "Compra"]
            .groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)": "Compra (MWh)"})
        )

        vendas_net = (
            base[base["Operação"] == "Venda"]
            .groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"]
            .sum()
            .rename(columns={"Volume (MWh)": "Venda (MWh)"})
        )

        nets = compras_net.merge(
            vendas_net,
            on=["Parte", "Contraparte", "Submercado", "Tipo de Energia"],
            how="inner"
        )

        if pagina == "Base Conferência":

            st.subheader("Base Conferência")

            base_exibicao = base.copy()
            base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            base_exibicao["Volume MWm"]   = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}")

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
            vendas_calc  = encontro[encontro["Operação"] == "Venda"]

            compras = compras_calc.copy()
            vendas  = vendas_calc.copy()

            compras["Volume (MWh)"] = compras["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            compras["Volume MWm"]   = compras["Volume MWm"].map(lambda x: f"{x:.6f}")
            vendas["Volume (MWh)"]  = vendas["Volume (MWh)"].map(lambda x: f"{x:.3f}")
            vendas["Volume MWm"]    = vendas["Volume MWm"].map(lambda x: f"{x:.6f}")

            st.markdown("## COMPRAS")
            st.dataframe(compras[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)

            st.markdown("## VENDAS")
            st.dataframe(vendas[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)

            total_compra     = compras_calc["Volume (MWh)"].sum()
            total_venda      = vendas_calc["Volume (MWh)"].sum()
            saldo            = total_compra - total_venda
            total_compra_mwm = compras_calc["Volume MWm"].sum()
            total_venda_mwm  = vendas_calc["Volume MWm"].sum()

            mes_referencia = int(df["Mes"].dropna().iloc[0])
            saldo_mwm = saldo / horas_mes.get(mes_referencia, 744)

            ajuste = contraparte if saldo > 0 else parte if saldo < 0 else "ZERADO"

            resumo = pd.DataFrame({
                "Tipo": ["Compras", "Vendas", "Saldo"],
                "MWh":  [f"{total_compra:.3f}", f"{total_venda:.3f}", f"{saldo:.3f}"],
                "MWm":  [f"{total_compra_mwm:.6f}", f"{total_venda_mwm:.6f}", f"{saldo_mwm:.6f}"]
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
