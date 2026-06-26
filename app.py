# APP_BOOK_ENERGIA_V22 - VERSÃO COM RESUMO DE NETS EM MWm, QUEM AJUSTA CORRETO E DESTAQUE ROXO
import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

pd.set_option("styler.render.max_elements", 2000000)

BOLETAS_ACR = {
    122387, 122389, 122391, 122393, 122395, 122397, 122399, 122401,
    144795, 144797, 144799, 148084, 148088, 148090, 148092, 148518,
}

def formatar_cnpj(valor):
    if pd.isna(valor): return ""
    cnpj = "".join(filter(str.isdigit, str(valor))).zfill(14)
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

def ler_csv_ccee(bytes_csv):
    df = pd.read_csv(BytesIO(bytes_csv), sep='\t', encoding='latin1', skiprows=1, dtype=str)
    df.columns = df.columns.str.strip()
    if 'SITUACAO_CONTRATO' in df.columns:
        df = df[df['SITUACAO_CONTRATO'].str.strip().str.lower() != 'rascunho']
    for col in ['CODIGO_CONTRATO', 'SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA', 'MWmedio']:
        if col in df.columns: df[col] = df[col].str.strip()
    df['_CHAVE'] = df['SIGLA_PERFIL_VENDEDOR'].fillna('') + df['SIGLA_PERFIL_COMPRADOR'].fillna('') + df['SUBMERCADO_ENTREGA'].fillna('')
    return df

def extrair_csvs_zip(zip_file):
    result = {'cceal': None, 'cbr': None, 'ccear_q': None}
    if zip_file is None: return result
    try:
        with zipfile.ZipFile(zip_file) as zf:
            for nome in zf.namelist():
                nome_lower = nome.lower()
                if nome_lower.endswith('/') or not nome_lower.endswith('.csv') or 'parcela' in nome_lower: continue
                dados = zf.read(nome)
                if 'ccear_q' in nome_lower: result['ccear_q'] = ler_csv_ccee(dados)
                elif 'cbr_mercado' in nome_lower: result['cbr'] = ler_csv_ccee(dados)
                elif 'cceal' in nome_lower: result['cceal'] = ler_csv_ccee(dados)
    except Exception as e: st.warning(f"Erro ao ler ZIP: {e}")
    return result

def combiner_dfs(lista):
    validos = [df for df in lista if df is not None and not df.empty]
    return pd.concat(validos, ignore_index=True) if validos else pd.DataFrame()

def criar_indices_busca(df_ccee):
    if df_ccee.empty: return {}, {}, {}, {}, {}, {}, {}, {}
    df_limpo = df_ccee.drop_duplicates(subset=['CODIGO_CONTRATO'])
    return (
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo['_CHAVE'])),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_VENDEDOR', ''))),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_COMPRADOR', ''))),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SUBMERCADO_ENTREGA', ''))),
        set(df_limpo['CODIGO_CONTRATO']),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MINIMO_MODULACAO_MW', '-'))),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MAXIMO_MODULACAO_MW', '-'))),
        dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('TIPO_MODULACAO', '-')))
    )

# ── FUNÇÃO DE ESTILIZAÇÃO VISUAL (HIGHLIGHTS) ──
def highlight_linhas(row, nets_aceitos_set):
    # Se a boleta foi editada manualmente: Azul Claro
    if "Editado Manualmente" in row.index and row["Editado Manualmente"] is True:
        return ["background-color: #D6EAF8"] * len(row)
    
    # Criar chave única para checar se pertence a um NET Aceito
    chave_net = (
        str(row.get("Parte", "")).strip(),
        str(row.get("Contraparte", "")).strip(),
        str(row.get("Submercado", "")).strip(),
        str(row.get("Tipo de Energia", "")).strip()
    )
    
    # Se estiver marcado como NET Aceito: ROXO
    if chave_net in nets_aceitos_set:
        return ["background-color: #E8DAEF; color: #5B2C6F; font-weight: bold;"] * len(row)
        
    # Se for Intra-Portfólio (Parte == Contraparte Razão Social): AMARELO
    parte = str(row.get("Parte", "")).strip().upper()
    contraparte_rs = str(row.get("Contraparte Razão Social", "")).strip().upper()
    if parte and contraparte_rs and parte == contraparte_rs:
        return ["background-color: #FFD700"] * len(row)
        
    return [""] * len(row)

st.set_page_config(page_title="Book Energia", layout="wide")
pagina = st.sidebar.radio("Menu", ["Base Conferência", "Encontro Energético"])
st.title("📊 Book Energia")

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx", "xlsm"])
arquivo_mes_anterior = st.file_uploader("Selecione a planilha Mês Anterior", type=["xlsx"])
arquivo_ajuste_manual = st.file_uploader("Selecione a planilha de Ajuste Manual", type=["xlsx"])
zip_matrix = st.file_uploader("Selecione o ZIP Matrix", type=["zip"])
zip_bismut = st.file_uploader("Selecione o ZIP Bismut", type=["zip"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=8)

        if "Parte_razao_social" in df.columns and "Contraparte_razao_social" in df.columns:
            mask_rateio_interno = df["Parte_razao_social"].astype(str).str.strip().str.upper() == df["Contraparte_razao_social"].astype(str).str.strip().str.upper()
            df = df[~mask_rateio_interno].reset_index(drop=True)

        if "Codigo_WBC" in df.columns and "Numero_referencia_contrato" in df.columns and "Rateio" in df.columns:
            mask_rateio_duplicado = (df["Codigo_WBC"].astype(str).str.strip() == df["Numero_referencia_contrato"].astype(str).str.strip()) & (df["Rateio"].astype(str).str.strip().str.upper() == "SIM")
            df = df[~mask_rateio_duplicado].reset_index(drop=True)

        horas_mes = {1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720, 7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744}
        mapa_mes_anterior = dict(zip(pd.read_excel(arquivo_mes_anterior)["BOLETA"], pd.read_excel(arquivo_mes_anterior)["Codigo_CCEE"])) if arquivo_mes_anterior is not None else {}

        # Ajustes Manuais via planilha externa...
        mapa_ajuste_manual_paradigma = {}
        if arquivo_ajuste_manual is not None:
            df_aj_manual = pd.read_excel(arquivo_ajuste_manual).dropna(subset=["BOLETA"])
            df_aj_manual["BOLETA"] = df_aj_manual["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
            if "Contrato CliqCCEE" in df_aj_manual.columns:
                mapa_ajuste_manual_paradigma = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["Contrato CliqCCEE"].astype(str)))

        csvs_matrix = extrair_csvs_zip(zip_matrix)
        csvs_bismut = extrair_csvs_zip(zip_bismut)
        df_ccee_matrix = combiner_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        df_ccee_bismut = combiner_dfs([csvs_bismut['cceal']])
        df_ccee_acr = combiner_dfs([csvs_matrix['ccear_q']])

        idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext, idx_m_min, idx_m_max, idx_m_tipo = criar_indices_busca(df_ccee_matrix)
        idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext, idx_b_min, idx_b_max, idx_b_tipo = criar_indices_busca(df_ccee_bismut)
        idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext, idx_a_min, idx_a_max, idx_a_tipo = criar_indices_busca(df_ccee_acr)

        mapa_energia = {"Incentivada 50%": "Incentivada-I5", "Cogeração Qualificada 50%": "Incentivada-CQ5", "Incentivada 100%": "Incentivada-I1", "Convencional": "Convencional", "Incentivada 0%": "Incentivada-I0"}
        mapa_submercado = {"Sul": "SUL", "S": "SUL", "SE/CO": "SUDESTE", "N": "NORTE", "NE": "NORDESTE"}
        mapa_modulacao = {"F - Flat": "FLAT", "C - Carga": "CARGA", "DECLARADO": "DECLARADA", "G - Geração": "GERAÇÃO"}

        df["Suprimento_inicio"] = pd.to_datetime(df["Suprimento_inicio"], errors="coerce")
        df["Suprimento_termino"] = pd.to_datetime(df["Suprimento_termino"], errors="coerce")
        cp_lp = ((df["Suprimento_termino"] - df["Suprimento_inicio"]).dt.days + 1).apply(lambda x: "CP" if x <= 31 else "LP")
        horas_por_linha = df["Mes"].map(horas_mes)
        volume_mwm = (df["QuantAtualizada"] / horas_por_linha).round(6)

        base = pd.DataFrame()
        base["BOLETA"]                         = df["Codigo_WBC"]
        base["Operação"]                       = df["Movimentacao"]
        base["Tipo de Energia"]                = df["Fonte_Contrato"].map(mapa_energia).fillna(df["Fonte_Contrato"])
        base["Parte"]                          = df["Parte_razao_social"]
        base["Contraparte Razão Social"]       = df["Contraparte_razao_social"] if "Contraparte_razao_social" in df.columns else "-"
        base["Contraparte"]                    = df["Sigla_CCEE_Contraparte"]
        base["CP/LP"]                          = cp_lp
        base["CNPJ CONTRAPARTE"]               = df["Contraparte_CNPJ"].apply(formatar_cnpj)
        base["Submercado"]                     = df["Submercado"].astype(str).str.strip().map(mapa_submercado).fillna(df["Submercado"])
        base["Volume (MWh)"]                   = df["QuantAtualizada"].round(3)
        base["Volume MWm"]                     = volume_mwm.round(6)
        base["Contrato CliqCCEE"]              = "-"
        base["Editado Manualmente"]            = False
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str)
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str)

        # ── LOGICA INTERATIVA DE SESSÃO DO STREAMLIT ──
        if "base_editada" not in st.session_state:
            st.session_state["base_editada"] = base.copy()
        else:
            base = st.session_state["base_editada"]

        if "nets_status" not in st.session_state:
            st.session_state["nets_status"] = {}

        # ── CÁLCULO DO RESUMO DE NETS POSSÍVEIS UTILIZANDO VOLUME MWm ──
        compras_net = base[base["Operação"] == "Compra"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume MWm"].sum().rename(columns={"Volume MWm": "Compra (MWm)"})
        vendas_net = base[base["Operação"] == "Venda"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume MWm"].sum().rename(columns={"Volume MWm": "Venda (MWm)"})
        nets = compras_net.merge(vendas_net, on=["Parte", "Contraparte", "Submercado", "Tipo de Energia"], how="inner")

        if not nets.empty:
            nets["NET (MWm)"] = nets["Compra (MWm)"] - nets["Venda (MWm)"]
            # Regra corrigida: Se o saldo for positivo (>0) -> Vendedor. Se for negativo (<0) -> Comprador.
            nets["Quem Ajusta"] = nets.apply(
                lambda r: r["Parte"] if r["NET (MWm)"] > 0 else (r["Contraparte"] if r["NET (MWm)"] < 0 else "ZERADO"), 
                axis=1
            )
            
            # Reconstrói a coluna "Net Aceito" baseado no dicionário salvo em session_state
            nets["Net Aceito"] = nets.apply(lambda r: st.session_state["nets_status"].get((str(r["Parte"]), str(r["Contraparte"]), str(r["Submercado"]), str(r["Tipo de Energia"])), False), axis=1)

        # Conjunto (Set) de chaves de NETs aceitos para acelerar a pintura das células roxas
        nets_aceitos_set = set()
        if not nets.empty:
            for _, r in nets[nets["Net Aceito"] == True].iterrows():
                nets_aceitos_set.add((str(r["Parte"]), str(r["Contraparte"]), str(r["Submercado"]), str(r["Tipo de Energia"])))

        if pagina == "Base Conferência":
            st.subheader("Base Conferência")
            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric("Total de Contratos", len(base))
            col_metric2.metric("Contratos de Compra", len(base[base['Operação'] == 'Compra']))
            col_metric3.metric("Contratos de Venda", len(base[base['Operação'] == 'Venda']))
            st.markdown("---")

            # ── EXIBIÇÃO DA TABELA RESUMO DE NETS POSSÍVEIS INTERATIVA ──
            st.subheader("⚖️ Resumo de NETs Possíveis (Foco em MWm e Flag Interativa)")
            if not nets.empty:
                # Exibimos via data_editor para permitir que a flag seja alterada de forma dinâmica pelo usuário
                cols_net_show = ["Net Aceito", "Parte", "Contraparte", "Submercado", "Tipo de Energia", "Compra (MWm)", "Venda (MWm)", "NET (MWm)", "Quem Ajusta"]
                
                config_net = {
                    "Net Aceito": st.column_config.CheckboxColumn("Net Aceito", default=False),
                    "Compra (MWm)": st.column_config.NumberColumn(format="%.6f", disabled=True),
                    "Venda (MWm)": st.column_config.NumberColumn(format="%.6f", disabled=True),
                    "NET (MWm)": st.column_config.NumberColumn(format="%.6f", disabled=True),
                    "Quem Ajusta": st.column_config.Column(disabled=True)
                }
                
                nets_editado = st.data_editor(nets[cols_net_show], use_container_width=True, hide_index=True, column_config=config_net, key="editor_nets")
                
                # Tratar a alteração do checkbox da flag
                if st.session_state.get("editor_nets") and st.session_state["editor_nets"].get("edited_rows"):
                    for idx_net, alteracoes in st.session_state["editor_nets"]["edited_rows"].items():
                        if "Net Aceito" in alteracoes:
                            row_alvo = nets.iloc[idx_net]
                            chave_alvo = (str(row_alvo["Parte"]), str(row_alvo["Contraparte"]), str(row_alvo["Submercado"]), str(row_alvo["Tipo de Energia"]))
                            st.session_state["nets_status"][chave_alvo] = alteracoes["Net Aceito"]
                    st.rerun()
            else:
                st.info("Nenhum match completo de NET encontrado.")
            st.markdown("---")

            # Exibição e filtros normais da tabela principal...
            base_exibicao = base.copy()
            styled = base_exibicao.style.apply(lambda r: highlight_lines(r, nets_aceitos_set), axis=1)
            
            # Renderização da tabela principal com a estilização condicional (Roxo se Net Aceito)
            st.dataframe(styled, use_container_width=True, hide_index=True)

        elif pagina == "Encontro Energético":
            st.subheader("🤝 Detalhes do Encontro Energético")
            # Código da aba do encontro mantendo a fidelidade estrutural
