# APP_BOOK_ENERGIA_V18
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V18: + Seção "Contratos sem Match"

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
                if nome_lower.endswith('/'):
                    continue
                if not nome_lower.endswith('.csv'):
                    continue
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


def combiner_dfs(lista):
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


def _selecionar_df_para_boleta(boleta, is_bismut, df_matrix, df_bismut, df_acr):
    """Retorna o DataFrame correto para a boleta, seguindo o mesmo roteamento de resolver_contrato_cliqccee."""
    try:
        boleta_int = int(float(str(boleta).strip()))
    except (ValueError, TypeError):
        boleta_int = -1

    if boleta_int in BOLETAS_ACR:
        return df_acr
    elif is_bismut:
        return df_bismut
    else:
        return df_matrix


def _buscar_linha_contrato(codigo, df_ccee):
    """
    Retorna a primeira linha não-RASCUNHO encontrada para o código, ou None.
    """
    if df_ccee.empty or pd.isna(codigo) or str(codigo).strip() in ('', '-', 'None'):
        return None
    codigo = str(codigo).strip()
    encontrado = df_ccee[df_ccee['CODIGO_CONTRATO'] == codigo]
    if encontrado.empty:
        return None
    for _, row in encontrado.iterrows():
        situacao = str(row.get('SITUACAO_CONTRATO', '')).strip().lower()
        if situacao != 'rascunho':
            return row
    return None


def verificar_contrato_sem_match(row, df_matrix, df_bismut, df_acr):
    """
    Verifica se a boleta possui correspondência válida no CSV.
    Retorna None se o contrato é válido, ou uma string de justificativa se há inconsistência.
    """
    BISMUT_NOME = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."
    is_bismut = str(row["Parte"]).strip().upper() == BISMUT_NOME.upper()

    codigos = [
        row.get("Contrato CliqCCEE", ""),
        row.get("Contrato CliqCCEE mês anterior", ""),
        row.get("CliqCCEE Paradigma", ""),
    ]

    # Ao menos um código deve estar preenchido para que a boleta seja avaliada
    tem_codigo = any(
        str(c).strip() not in ('', '-', 'None', 'nan')
        for c in codigos
    )
    if not tem_codigo:
        return None

    df = _selecionar_df_para_boleta(row["BOLETA"], is_bismut, df_matrix, df_bismut, df_acr)

    vendedor_boleta   = str(row["Vendedor"]).strip()
    comprador_boleta  = str(row["Comprador"]).strip()
    submercado_boleta = str(row["Submercado"]).strip()

    linha_encontrada = None
    for codigo in codigos:
        linha = _buscar_linha_contrato(codigo, df)
        if linha is not None:
            linha_encontrada = linha
            break

    if linha_encontrada is None:
        return "Contrato inexistente"

    # Contrato encontrado — verificar divergências
    vendedor_csv   = str(linha_encontrada.get('SIGLA_PERFIL_VENDEDOR', '')).strip()
    comprador_csv  = str(linha_encontrada.get('SIGLA_PERFIL_COMPRADOR', '')).strip()
    submercado_csv = str(linha_encontrada.get('SUBMERCADO_ENTREGA', '')).strip()

    div_vendedor   = vendedor_boleta   != vendedor_csv
    div_comprador  = comprador_boleta  != comprador_csv
    div_submercado = submercado_boleta != submercado_csv

    if not div_vendedor and not div_comprador and not div_submercado:
        return None  # match perfeito

    partes = []
    if div_vendedor:
        partes.append("Vendedor")
    if div_comprador:
        partes.append("Comprador")
    if div_submercado:
        partes.append(f"Submercado (Boleta={submercado_boleta} | CSV={submercado_csv})")

    if len(partes) == 1:
        return f"Divergência de {partes[0]}"
    elif len(partes) == 2:
        return f"Divergência de {partes[0]} e {partes[1]}"
    else:
        return f"Divergência de {partes[0]}, {partes[1]} e {partes[2]}"


def highlight_mesmo_titular(row):
    """
    Pinta a linha de amarelo quando Parte == Contraparte Razão Social.
    """
    if "Editado Manualmente" in row.index and row["Editado Manualmente"] is True:
        return ["background-color: #D6EAF8"] * len(row)
    parte = str(row.get("Parte", "")).strip().upper()
    contraparte_rs = str(row.get("Contraparte Razão Social", "")).strip().upper()
    if parte and contraparte_rs and parte == contraparte_rs:
        return ["background-color: #FFD700"] * len(row)
    return [""] * len(row)


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
        df_ccee_matrix = combiner_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        # DataFrame Bismut: cceal_firme do ZIP Bismut
        df_ccee_bismut = combiner_dfs([csvs_bismut['cceal']])
        # DataFrame ACR: ccear_q do ZIP Matrix
        df_ccee_acr = combiner_dfs([csvs_matrix['ccear_q']])

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
        # ── NOVO: Contraparte Razão Social logo após Parte ─────────────────────
        base["Contraparte Razão Social"]       = df["Contraparte_razao_social"] if "Contraparte_razao_social" in df.columns else "-"
        # ───────────────────────────────────────────────────────────────────────
        base["Contraparte"]                    = df["Sigla_CCEE_Contraparte"].fillna("-").astype(str)
        base["CP/LP"]                          = cp_lp
        base["CNPJ CONTRAPARTE"]               = df["Contraparte_CNPJ"].apply(formatar_cnpj)
        base["Submercado"]                     = df["Submercado"].astype(str).str.strip().map(mapa_submercado).fillna(df["Submercado"])
        base["Volume (MWh)"]                   = df["QuantAtualizada"].round(3)
        base["Volume MWm"]                     = volume_mwm.round(6)
        base["CliqCCEE Paradigma"]             = df["Codigo_CCEE"].fillna("-").astype(str)
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["Modulação Mínima"]               = df["FlexLimite_modulacaoMin"].fillna("-")
        base["Modulação Máxima"]               = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-").astype(str)
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str)
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str)

        # ── Flag: zera volumes quando Parte == Contraparte Razão Social ────────
        mask_mesmo_titular = (
            base["Parte"].astype(str).str.strip().str.upper()
            == base["Contraparte Razão Social"].astype(str).str.strip().str.upper()
        )
        base.loc[mask_mesmo_titular, "Volume (MWh)"] = 0.0
        base.loc[mask_mesmo_titular, "Volume MWm"]   = 0.0
        # ───────────────────────────────────────────────────────────────────────

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
            base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee, axis=1).astype(str)
        else:
            base["Contrato CliqCCEE"] = "-"
            if pagina == "Base Conferência":
                st.info("ℹ️ Faça upload dos ZIPs para preencher a coluna 'Contrato CliqCCEE'.")
        # ───────────────────────────────────────────────────────────────────────

        base["Editado Manualmente"] = False

        if "base_editada" not in st.session_state:
            st.session_state["base_editada"] = base.copy()
        else:
            # Garantir que se a planilha ou zips mudarem, linhas novas sejam sincronizadas, mantendo edições antigas indexadas por BOLETA
            df_atual = base.copy()
            df_salvo = st.session_state["base_editada"]
            df_salvo = df_salvo[df_salvo["Editado Manualmente"] == True]
            if not df_salvo.empty:
                df_atual.set_index("BOLETA", inplace=True)
                df_salvo.set_index("BOLETA", inplace=True)
                df_atual.update(df_salvo)
                df_atual.reset_index(inplace=True)
                # Recalcular contratos que foram afetados por mudanças em outras colunas mas NÃO foram editados diretamente no contrato cliqccee
                if csvs_disponiveis:
                    linhas_para_recalcular = df_atual["Editado Manualmente"] & (~df_atual["BOLETA"].isin(st.session_state.get("contratos_editados_diretamente", [])))
                    if linhas_para_recalcular.any():
                        df_atual.loc[linhas_para_recalcular, "Contrato CliqCCEE"] = df_atual[linhas_para_recalcular].apply(calcular_contrato_cliqccee, axis=1).astype(str)
                base = df_atual
            else:
                base = df_atual
            st.session_state["base_editada"] = base.copy()

        base = st.session_state["base_editada"]

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

            # ── Indicadores Visuais de Contratos (Adicionados) ──────────────────
            total_contratos = len(base)
            total_compras = len(base[base['Operação'].str.upper() == 'COMPRA'])
            total_vendas = len(base[base['Operação'].str.upper() == 'VENDA'])

            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric(label="Total de Contratos", value=total_contratos)
            col_metric2.metric(label="Contratos de Compra 📥", value=total_compras)
            col_metric3.metric(label="Contratos de Venda 📤", value=total_vendas)

            st.markdown("---")
            # ───────────────────────────────────────────────────────────────────

            # ── Flags da Base Conferência ──────────────────────────────────────
            col_flag1, col_flag2 = st.columns(2)
            with col_flag1:
                flag_mesmo_titular = st.toggle(
                    "🟡 Ocultar IntraPortifólio",
                    value=True
                )
            with col_flag2:
                flag_ocultar_zerados = st.toggle(
                    "🚫 Ocultar contratos zerados (Volume MWh = 0)",
                    value=False
                )
            # ───────────────────────────────────────────────────────────────────

            base_exibicao = base.copy()

            st.markdown("### 🔎 Filtros")

            col_f1, col_f2, col_f3 = st.columns(3)

            with col_f1:
                filtro_operacao = st.multiselect(
                    "Operação",
                    options=sorted(base_exibicao["Operação"].dropna().unique()),
                    default=[]
                )

            with col_f2:
                filtro_status = st.multiselect(
                    "Contrato CliqCCEE",
                    options=sorted(base_exibicao["Contrato CliqCCEE"].dropna().astype(str).unique()),
                    default=[]
                )

            with col_f3:
                filtro_submercado = st.multiselect(
                    "Submercado",
                    options=sorted(base_exibicao["Submercado"].dropna().astype(str).unique()),
                    default=[]
                )

            col_f4, col_f5, col_f6 = st.columns(3)

            with col_f4:
                filtro_parte = st.text_input("Parte")

            with col_f5:
                filtro_contraparte = st.text_input("Contraparte")

            with col_f6:
                filtro_boleta = st.text_input("Boleta")

            if filtro_operacao:
                base_exibicao = base_exibicao[base_exibicao["Operação"].isin(filtro_operacao)]

            if filtro_status:
                base_exibicao = base_exibicao[base_exibicao["Contrato CliqCCEE"].astype(str).isin(filtro_status)]

            if filtro_submercado:
                base_exibicao = base_exibicao[base_exibicao["Submercado"].astype(str).isin(filtro_submercado)]

            if filtro_parte:
                base_exibicao = base_exibicao[
                    base_exibicao["Parte"].astype(str).str.contains(filtro_parte, case=False, na=False)
                ]

            if filtro_contraparte:
                base_exibicao = base_exibicao[
                    base_exibicao["Contraparte"].astype(str).str.contains(filtro_contraparte, case=False, na=False)
                ]

            if filtro_boleta:
                base_exibicao = base_exibicao[
                    base_exibicao["BOLETA"].astype(str).str.contains(filtro_boleta, case=False, na=False)
                ]

            if flag_ocultar_zerados:
                base_exibicao = base_exibicao[base_exibicao["Volume (MWh)"] != 0.0]

            base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            base_exibicao["Volume MWm"]   = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

            st.caption(f"{len(base_exibicao):,} registros encontrados")

            col_config = {
                "BOLETA": st.column_config.Column(disabled=True),
                "Operação": st.column_config.Column(disabled=True),
                "Tipo de Energia": st.column_config.Column(disabled=True),
                "Parte": st.column_config.Column(disabled=True),
                "Contraparte Razão Social": st.column_config.Column(disabled=True),
                "Contraparte": st.column_config.TextColumn(disabled=False),
                "CP/LP": st.column_config.Column(disabled=True),
                "CNPJ CONTRAPARTE": st.column_config.Column(disabled=True),
                "Submercado": st.column_config.Column(disabled=True),
                "Volume (MWh)": st.column_config.Column(disabled=True),
                "Volume MWm": st.column_config.Column(disabled=True),
                "CliqCCEE Paradigma": st.column_config.TextColumn(disabled=False),
                "Modulação WBC": st.column_config.Column(disabled=True),
                "Modulação Mínima": st.column_config.Column(disabled=True),
                "Modulação Máxima": st.column_config.Column(disabled=True),
                "Contrato CliqCCEE mês anterior": st.column_config.Column(disabled=True),
                "Vendedor": st.column_config.Column(disabled=True),
                "Comprador": st.column_config.Column(disabled=True),
                "Contrato CliqCCEE": st.column_config.Column(disabled=True),
                "Editado Manualmente": st.column_config.Column(disabled=True),
            }

            if flag_mesmo_titular:
                styled = base_exibicao.style.apply(highlight_mesmo_titular, axis=1)
                base_editada_df = st.data_editor(styled, use_container_width=True, hide_index=True, column_config=col_config, key="editor_base")
            else:
                base_editada_df = st.data_editor(base_exibicao, use_container_width=True, hide_index=True, column_config=col_config, key="editor_base")

            if st.session_state.get("editor_base") and st.session_state["editor_base"].get("edited_rows"):
                edicoes = st.session_state["editor_base"]["edited_rows"]
                indices_exibicao = base_exibicao.index.tolist()
                
                if "contratos_editados_diretamente" not in st.session_state:
                    st.session_state["contratos_editados_diretamente"] = []

                for idx_str, alteracoes in edicoes.items():
                    idx = int(idx_str)
                    idx_real = indices_exibicao[idx]
                    boleta_alvo = base.loc[idx_real, "BOLETA"]
                    
                    base.loc[idx_real, "Editado Manualmente"] = True
                    for col, val in alteracoes.items():
                        base.loc[idx_real, col] = str(val)
                        if col == "Contrato CliqCCEE":
                            if boleta_alvo not in st.session_state["contratos_editados_diretamente"]:
                                st.session_state["contratos_editados_diretamente"].append(boleta_alvo)
                    
                    if csvs_disponiveis and "Contrato CliqCCEE" not in alteracoes:
                        if boleta_alvo not in st.session_state["contratos_editados_diretamente"]:
                            novo_contrato = calcular_contrato_cliqccee(base.loc[idx_real])
                            base.loc[idx_real, "Contrato CliqCCEE"] = str(novo_contrato)
                
                st.session_state["base_editada"] = base.copy()
                st.rerun()

            # Download sempre com dados numéricos originais (sem formatação de string)
            base_download = base.copy()
            if flag_ocultar_zerados:
                base_download = base_download[base_download["Volume (MWh)"] != 0.0]

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                base_download.to_excel(writer, sheet_name="Base Conferência", index=False)

            st.download_button(
                "📥 Download Base Conferência",
                data=output.getvalue(),
                file_name="Base_Conferencia.xlsx"
            )

            # ── Contratos sem Match (inline, abaixo da tabela principal) ──────
            if csvs_disponiveis:
                st.markdown("---")
                st.subheader("Contratos sem Match")

                resultados = []
                for _, row in base.iterrows():
                    # Ignorar contratos zerados
                    if float(row["Volume (MWh)"]) == 0.0:
                        continue
                    justificativa = verificar_contrato_sem_match(
                        row,
                        df_matrix=df_ccee_matrix,
                        df_bismut=df_ccee_bismut,
                        df_acr=df_ccee_acr,
                    )
                    if justificativa is not None:
                        resultados.append({
                            "Boleta":        row["BOLETA"],
                            "Vendedor":      row["Vendedor"],
                            "Comprador":     row["Comprador"],
                            "Justificativa": justificativa,
                        })

                df_sem_match = pd.DataFrame(resultados, columns=["Boleta", "Vendedor", "Comprador", "Justificativa"])

                st.caption(f"{len(df_sem_match):,} contrato(s) sem match encontrado(s)")
                st.dataframe(df_sem_match, use_container_width=True, hide_index=True)

                output_sm = BytesIO()
                with pd.ExcelWriter(output_sm, engine="openpyxl") as writer:
                    df_sem_match.to_excel(writer, sheet_name="Contratos sem Match", index=False)

                st.download_button(
                    "📥 Download Contratos sem Match",
                    data=output_sm.getvalue(),
                    file_name="Contratos_sem_Match.xlsx"
                )
            # ───────────────────────────────────────────────────────────────────

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

            compras["Volume (MWh)"] = compras["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            compras["Volume MWm"]   = compras["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            vendas["Volume (MWh)"]  = vendas["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            vendas["Volume MWm"]    = vendas["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

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
