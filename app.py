# APP_BOOK_ENERGIA_V22 - CORREÇÃO DE CHAVES DE BUSCA NET (SIGLAS CCEE)
import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

# Configura o limite do Pandas Styler para evitar o erro de estouro de células
pd.set_option("styler.render.max_elements", 2000000)

BOLETAS_ACR = {
    122387, 122389, 122391, 122393, 122395, 122397, 122399, 122401,
    144795, 144797, 144799, 148084, 148088, 148090, 148092, 148518,
}

def formatar_cnpj(valor):
    if pd.isna(valor):
        return ""
    cnpj = "".join(filter(str.isdigit, str(valor)))
    cnpj = cnpj.zfill(14)
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

def ler_csv_ccee(bytes_csv):
    df = pd.read_csv(BytesIO(bytes_csv), sep='\t', encoding='latin1', skiprows=1, dtype=str)
    df.columns = df.columns.str.strip()
    if 'SITUACAO_CONTRATO' in df.columns:
        df = df[df['SITUACAO_CONTRATO'].str.strip().str.lower() != 'rascunho']
    for col in ['CODIGO_CONTRATO', 'SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA', 'MWmedio']:
        if col in df.columns:
            df[col] = df[col].str.strip()
    df['_CHAVE'] = df['SIGLA_PERFIL_VENDEDOR'].fillna('') + df['SIGLA_PERFIL_COMPRADOR'].fillna('') + df['SUBMERCADO_ENTREGA'].fillna('')
    return df

def extrair_csvs_zip(zip_file):
    result = {'cceal': None, 'cbr': None, 'ccear_q': None}
    if zip_file is None:
        return result
    try:
        with zipfile.ZipFile(zip_file) as zf:
            for nome in zf.namelist():
                nome_lower = nome.lower()
                if nome_lower.endswith('/') or not nome_lower.endswith('.csv') or 'parcela' in nome_lower:
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
    validos = [df for df in lista if df is not None and not df.empty]
    if not validos:
        return pd.DataFrame()
    return pd.concat(validos, ignore_index=True)

def criar_indices_busca(df_ccee):
    if df_ccee.empty:
        return {}, {}, {}, {}, {}, {}, {}, {}
    df_limpo = df_ccee.drop_duplicates(subset=['CODIGO_CONTRATO'])
    dict_chave = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo['_CHAVE']))
    dict_vend = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_VENDEDOR', '')))
    dict_comp = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_COMPRADOR', '')))
    dict_sub = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SUBMERCADO_ENTREGA', '')))
    dict_lim_min = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MINIMO_MODULACAO_MW', '-')))
    dict_lim_max = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MAXIMO_MODULACAO_MW', '-')))
    dict_tipo_mod = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('TIPO_MODULACAO', '-')))
    set_existentes = set(df_limpo['CODIGO_CONTRATO'])
    return dict_chave, dict_vend, dict_comp, dict_sub, set_existentes, dict_lim_min, dict_lim_max, dict_tipo_mod

def calcular_nets(base_df, horas_mes):
    compras = base_df[(base_df["Operação"] == "Compra") & (base_df["Volume (MWh)"] > 0)].copy()
    vendas = base_df[(base_df["Operação"] == "Venda") & (base_df["Volume (MWh)"] > 0)].copy()
    if compras.empty or vendas.empty:
        return pd.DataFrame()
    
    # Agrupamos incluindo as colunas de siglas CCEE (Parte e Contraparte) e a nova coluna "Razão Social"
    compras_agg = compras.groupby(["Razão Social", "Parte", "Contraparte", "Submercado"], as_index=False)[["Volume (MWh)", "Volume MWm"]].sum()
    compras_agg.rename(columns={"Volume (MWh)": "Compra_MWh", "Volume MWm": "Compra_MWm"}, inplace=True)
    
    vendas_agg = vendas.groupby(["Razão Social", "Parte", "Contraparte", "Submercado"], as_index=False)[["Volume (MWh)", "Volume MWm"]].sum()
    vendas_agg.rename(columns={"Volume (MWh)": "Venda_MWh", "Volume MWm": "Venda_MWm"}, inplace=True)
    
    nets = compras_agg.merge(vendas_agg, on=["Razão Social", "Parte", "Contraparte", "Submercado"], how="inner")
    if nets.empty:
        return pd.DataFrame()
    
    nets["Saldo_MWh"] = nets["Compra_MWh"] - nets["Venda_MWh"]
    nets["Saldo_MWm"] = nets["Saldo_MWh"] / pd.Series([horas_mes.get(1, 744)] * len(nets))
    nets["Ajuste Net"] = nets.apply(lambda r: r["Contraparte"] if r["Saldo_MWm"] > 0 else (r["Parte"] if r["Saldo_MWm"] < 0 else "ZERADO"), axis=1)
    return nets

# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Book Energia", layout="wide")
pagina = st.sidebar.radio("Menu", ["Base Conferência", "Encontro Energético", "Arquivos CCEE"])
st.title("📊 Book Energia")

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx", "xlsm"])
arquivo_mes_anterior = st.file_uploader("Selecione a planilha Mês Anterior", type=["xlsx"])
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

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
            mapa_mes_anterior = dict(zip(df_mes_anterior["BOLETA"], df_mes_anterior["Codigo_CCEE"]))
        else:
            mapa_mes_anterior = {}

        csvs_matrix = extrair_csvs_zip(zip_matrix)
        csvs_bismut = extrair_csvs_zip(zip_bismut)

        df_ccee_matrix = combiner_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        df_ccee_bismut = combiner_dfs([csvs_bismut['cceal']])
        df_ccee_acr = combiner_dfs([csvs_matrix['ccear_q']])
        
        df_ccee_completo = combiner_dfs([df_ccee_matrix, df_ccee_bismut, df_ccee_acr])
        if not df_ccee_completo.empty and 'MWmedio' in df_ccee_completo.columns:
            df_ccee_completo['SIGLA_PERFIL_VENDEDOR'] = df_ccee_completo['SIGLA_PERFIL_VENDEDOR'].astype(str).str.strip()
            df_ccee_completo['SIGLA_PERFIL_COMPRADOR'] = df_ccee_completo['SIGLA_PERFIL_COMPRADOR'].astype(str).str.strip()
            df_ccee_completo['SUBMERCADO_ENTREGA'] = df_ccee_completo['SUBMERCADO_ENTREGA'].astype(str).str.strip()
            df_ccee_completo['MWmedio_num'] = df_ccee_completo['MWmedio'].astype(str).str.strip().str.replace(',', '.', regex=False)
            df_ccee_completo['MWmedio_num'] = pd.to_numeric(df_ccee_completo['MWmedio_num'], errors='coerce').fillna(0.0)

        idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext, idx_m_min, idx_m_max, idx_m_tipo = criar_indices_busca(df_ccee_matrix)
        idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext, idx_b_min, idx_b_max, idx_b_tipo = criar_indices_busca(df_ccee_bismut)
        idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext, idx_a_min, idx_a_max, idx_a_tipo = criar_indices_busca(df_ccee_acr)

        mapa_energia = {"Incentivada 50%": "Incentivada-I5", "Cogeração Qualificada 50%": "Incentivada-CQ5", "Incentivada 100%": "Incentivada-I1", "Convencional": "Convencional", "Incentivada 0%": "Incentivada-I0"}
        mapa_submercado = {"Sul": "SUL", "S": "SUL", "SE/CO": "SUDESTE", "N": "NORTE", "NE": "NORDESTE"}
        mapa_modulacao = {"F - Flat": "FLAT", "C - Carga": "CARGA", "DECLARADO": "DECLARADA", "G - Geração": "GERAÇÃO"}

        df["Suprimento_inicio"] = pd.to_datetime(df["Suprimento_inicio"], errors="coerce")
        df["Suprimento_termino"] = pd.to_datetime(df["Suprimento_termino"], errors="coerce")

        dias_periodo = (df["Suprimento_termino"] - df["Suprimento_inicio"]).dt.days + 1
        cp_lp = dias_periodo.apply(lambda x: "CP" if x <= 31 else "LP")
        horas_por_linha = df["Mes"].map(horas_mes)
        volume_mwm = (df["QuantAtualizada"] / horas_por_linha).round(6)

        codigo_ccee_str = df["Codigo_CCEE"].fillna("").astype(str).str.strip().replace("nan", "-").replace("", "-")
        codigo_ccee_str = codigo_ccee_str.apply(lambda x: "-" if x == "0" or x == "" else x)

        base = pd.DataFrame()
        base["BOLETA"]                         = df["Codigo_WBC"]
        base["Operação"]                       = df["Movimentacao"]
        base["Tipo de Energia"]                = df["Fonte_Contrato"].map(mapa_energia).fillna(df["Fonte_Contrato"])
        base["Razão Social"]                   = df["Parte_razao_social"]
        base["Contraparte Razão Social"]       = df["Contraparte_razao_social"] if "Contraparte_razao_social" in df.columns else "-"
        # Nova mapeação: Mapeia o Perfil_CCEE_Parte como a sigla curta da Parte
        base["Parte"]                          = df["Perfil_CCEE_Parte"].fillna("-").astype(str).str.strip()
        base["Contraparte"]                    = df["Sigla_CCEE_Contraparte"].fillna("-").astype(str).str.strip()
        base["CP/LP"]                          = cp_lp
        base["CNPJ CONTRAPARTE"]               = df["Contraparte_CNPJ"].apply(formatar_cnpj)
        base["Submercado"]                     = df["Submercado"].astype(str).str.strip().map(mapa_submercado).fillna(df["Submercado"])
        base["Volume (MWh)"]                   = df["QuantAtualizada"].round(3)
        base["Volume MWm"]                     = volume_mwm.round(6)
        base["CliqCCEE Paradigma"]             = codigo_ccee_str
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["% Modulação Mínima"]             = df["FlexLimite_modulacaoMin"].fillna("-")
        base["% Modulação Máxima"]             = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-").astype(str)
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str).str.strip()
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str).str.strip()
        base["Contrato CliqCCEE"]              = "-"

        csvs_disponiveis = any([not df_ccee_matrix.empty, not df_ccee_bismut.empty, not df_ccee_acr.empty])

        if csvs_disponiveis:
            BISMUT_NOME_UPPER = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."
            def calcular_contrato_cliqccee_fast(row):
                try: b_int = int(float(str(row["BOLETA"]).strip()))
                except: b_int = -1
                if b_int in BOLETAS_ACR: d_ch, s_ext = idx_a_chave, set_a_ext
                elif str(row["Razão Social"]).strip().upper() == BISMUT_NOME_UPPER: d_ch, s_ext = idx_b_chave, set_b_ext
                else: d_ch, s_ext = idx_m_chave, set_m_ext
                chave_esp = str(row["Vendedor"]).strip() + str(row["Comprador"]).strip() + str(row["Submercado"]).strip()
                c_ant = str(row["Contrato CliqCCEE mês anterior"]).strip()
                if c_ant in s_ext: return c_ant if d_ch.get(c_ant) == chave_esp else 'Verificar'
                c_par = str(row["CliqCCEE Paradigma"]).strip()
                if c_par in s_ext: return c_par if d_ch.get(c_par) == chave_esp else 'Verificar'
                return '-'
            base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee_fast, axis=1).astype(str)

        _vol_mwm_num = pd.to_numeric(base["Volume MWm"], errors="coerce")
        _mask_valido_book = _vol_mwm_num.notna() & (base["Volume MWm"].astype(str).str.strip() != "-")
        _df_book = base[["Contrato CliqCCEE"]].copy()
        _df_book["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
        base["Volume Book"] = _df_book.groupby("Contrato CliqCCEE")["_vol_num"].transform("sum")

        _vol_book_num = pd.to_numeric(base["Volume Book"], errors="coerce").fillna(0.0)
        _num_mod_min = pd.to_numeric(base["% Modulação Mínima"], errors="coerce").fillna(0.0)
        _num_mod_max = pd.to_numeric(base["% Modulação Máxima"], errors="coerce").fillna(0.0)
        _mask_tem_contrato = ~base["Contrato CliqCCEE"].astype(str).str.strip().isin(["", "-", "None", "nan", "Verificar"])
        base["Modulação Mínima"] = (_vol_book_num * (1 - (_num_mod_min / 100))).where(_mask_tem_contrato & (_num_mod_min > 0.0), "-")
        base["Modulação Máxima"] = (_vol_book_num * (1 + (_num_mod_max / 100))).where(_mask_tem_contrato & (_num_mod_max > 0.0), "-")

        if csvs_disponiveis:
            def buscar_campo_ccee(row, dict_m, dict_b, dict_a):
                cod = str(row["Contrato CliqCCEE"]).strip()
                if cod in ["", "-", "None", "nan", "Verificar"]: return "-"
                try: b_int = int(float(str(row["BOLETA"]).strip()))
                except: b_int = -1
                d_field = dict_a if b_int in BOLETAS_ACR else (dict_b if str(row["Razão Social"]).strip().upper() == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A." else dict_m)
                return d_field.get(cod, "-")
            base["Modulação Mínima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_min, idx_b_min, idx_a_min), axis=1)
            base["Modulação Máxima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_max, idx_b_max, idx_a_max), axis=1)
            base["Modulação CCEE"]        = base.apply(lambda r: buscar_campo_ccee(r, idx_m_tipo, idx_b_tipo, idx_a_tipo), axis=1)
        else:
            base["Modulação Mínima CCEE"] = base["Modulação Máxima CCEE"] = base["Modulação CCEE"] = "-"

        base["Check Modulação Mínima"] = base["Check Modulação Máxima"] = base["Check Modulação"] = "-"

        if csvs_disponiveis:
            _lista_dfs_ccee_vol = []
            for _df_src in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
                if _df_src is not None and not _df_src.empty and "CODIGO_CONTRATO" in _df_src.columns and "MWmedio" in _df_src.columns:
                    _tmp = _df_src[["CODIGO_CONTRATO", "MWmedio"]].copy()
                    _tmp["MWmedio"] = _tmp["MWmedio"].astype(str).str.strip().str.replace(",", ".", regex=False)
                    _tmp["MWmedio"] = pd.to_numeric(_tmp["MWmedio"], errors="coerce").fillna(0.0)
                    _lista_dfs_ccee_vol.append(_tmp)
            base["Volume CCEE"] = base["Contrato CliqCCEE"].map(pd.concat(_lista_dfs_ccee_vol, ignore_index=True).groupby("CODIGO_CONTRATO")["MWmedio"].sum() if _lista_dfs_ccee_vol else pd.Series()).fillna(0.0)
        else:
            base["Volume CCEE"] = 0.0

        _tol = 1e-6
        _diff_vol = pd.to_numeric(base["Volume Book"], errors="coerce").fillna(0.0) - pd.to_numeric(base["Volume CCEE"], errors="coerce").fillna(0.0)
        base["Check Volume"] = "OK"
        base.loc[_diff_vol > _tol, "Check Volume"] = "Book maior"
        base.loc[_diff_vol < -_tol, "Check Volume"] = "CCEE maior"

        # GERAÇÃO DE NETs DINÂMICO USANDO AS DUAS SIGLAS CURTAS DA CCEE
        df_nets = calcular_nets(base, horas_mes)
        
        if "nets_aceitos" not in st.session_state:
            st.session_state.nets_aceitos = {}
        
        if not df_nets.empty:
            vol_compra_cliq_lista = []
            vol_venda_cliq_lista = []
            check_nets_lista = []
            
            for _, net_row in df_nets.iterrows():
                sigla_parte = str(net_row["Parte"]).strip()
                sigla_contraparte = str(net_row["Contraparte"]).strip()
                submercado_net = str(net_row["Submercado"]).strip()
                saldo_esperado = abs(net_row["Saldo_MWm"])
                
                if not df_ccee_completo.empty:
                    # Volume Compra Cliq -> Contraparte vende para a Parte
                    v_compra_ccee = df_ccee_completo[
                        (df_ccee_completo["SIGLA_PERFIL_VENDEDOR"] == sigla_contraparte) & 
                        (df_ccee_completo["SIGLA_PERFIL_COMPRADOR"] == sigla_parte) &
                        (df_ccee_completo["SUBMERCADO_ENTREGA"] == submercado_net)
                    ]["MWmedio_num"].sum()
                    
                    # Volume Venda Cliq -> Parte vende para a Contraparte
                    v_venda_ccee = df_ccee_completo[
                        (df_ccee_completo["SIGLA_PERFIL_VENDEDOR"] == sigla_parte) & 
                        (df_ccee_completo["SIGLA_PERFIL_COMPRADOR"] == sigla_contraparte) &
                        (df_ccee_completo["SUBMERCADO_ENTREGA"] == submercado_net)
                    ]["MWmedio_num"].sum()
                else:
                    v_compra_ccee = 0.0
                    v_venda_ccee = 0.0
                
                vol_compra_cliq_lista.append(v_compra_ccee)
                vol_venda_cliq_lista.append(v_venda_ccee)
                
                ajuste_resp = net_row["Ajuste Net"]
                if ajuste_resp == "ZERADO":
                    check_nets_lista.append("-")
                else:
                    vol_efetivo_cliq = v_compra_ccee if ajuste_resp == sigla_parte else v_venda_ccee
                    diff_net = vol_efetivo_cliq - saldo_esperado
                    if abs(diff_net) < 1e-4: check_nets_lista.append("OK")
                    elif diff_net > 1e-4: check_nets_lista.append("Volume ajustado maior que o esperado")
                    else: check_nets_lista.append("Volume ajustado menor que o esperado")
            
            df_nets["Volume Compra Cliq"] = vol_compra_cliq_lista
            df_nets["Volume Venda Cliq"] = vol_venda_cliq_lista
            df_nets["Check Net"] = check_nets_lista

        # ──────────────────────────────────────────────────────────────────────────────
        if pagina == "Base Conferência":
            st.subheader("Base Conferência")
            
            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric(label="Total de Contratos", value=len(base))
            col_metric2.metric(label="Contratos de Compra 📥", value=len(base[base['Operação'].str.upper() == 'COMPRA']))
            col_metric3.metric(label="Contratos de Venda 📤", value=len(base[base['Operação'].str.upper() == 'VENDA']))
            st.markdown("---")

            # RESUMO DE NETs RETRÁTIL
            if not df_nets.empty:
                with st.expander("📋 Resumo de NETs (Clique para Expandir / Recolher)", expanded=True):
                    df_nets_display = df_nets[[
                        "Razão Social", "Parte", "Contraparte", "Submercado", "Compra_MWm", "Venda_MWm", 
                        "Saldo_MWm", "Ajuste Net", "Volume Compra Cliq", "Volume Venda Cliq", "Check Net"
                    ]].copy()
                    
                    df_nets_display.rename(columns={
                        "Compra_MWm": "Volume Total Compras (MWm)",
                        "Venda_MWm": "Volume Total Vendas (MWm)",
                        "Saldo_MWm": "Volume Net (MWm)"
                    }, inplace=True)
                    
                    for col in ["Volume Total Compras (MWm)", "Volume Total Vendas (MWm)", "Volume Net (MWm)", "Volume Compra Cliq", "Volume Venda Cliq"]:
                        df_nets_display[col] = df_nets_display[col].apply(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
                    
                    cols_net = st.columns(len(df_nets_display.columns) + 1)
                    cols_net[0].write("**Net Aceito**")
                    for idx, col_name in enumerate(df_nets_display.columns):
                        cols_net[idx + 1].write(f"**{col_name}**")
                    
                    for idx, (_, net_row) in enumerate(df_nets.iterrows()):
                        chave_net = (net_row["Razão Social"], net_row["Parte"], net_row["Contraparte"], net_row["Submercado"])
                        cols_data = st.columns(len(df_nets_display.columns) + 1)
                        
                        with cols_data[0]:
                            st.session_state.nets_aceitos[chave_net] = st.checkbox(
                                "Aceito", value=st.session_state.nets_aceitos.get(chave_net, False), key=f"net_aceito_{idx}", label_visibility="collapsed"
                            )
                        
                        for col_idx, col_value in enumerate(df_nets_display.iloc[idx]):
                            cols_data[col_idx + 1].write(str(col_value))
                st.markdown("---")

            # TABELA PRINCIPAL RETRÁTIL
            with st.expander("🔎 Filtros e Visão Global dos Contratos", expanded=True):
                col_flag1, col_flag2 = st.columns(2)
                with col_flag1: flag_mesmo_titular = st.toggle("Ocultar IntraPortifólio Visualmente", value=True)
                with col_flag2: flag_ocultar_zerados = st.toggle("Ocultar contratos zerados (Volume MWh = 0)", value=False)

                base_exibicao = base.copy()
                col_f1, col_f2, col_f3 = st.columns(3)
                filtro_operacao = col_f1.multiselect("Operação", options=sorted(base_exibicao["Operação"].dropna().unique()))
                filtro_status = col_f2.multiselect("Contrato CliqCCEE", options=sorted(base_exibicao["Contrato CliqCCEE"].dropna().astype(str).unique()))
                filtro_submercado = col_f3.multiselect("Submercado", options=sorted(base_exibicao["Submercado"].dropna().astype(str).unique()))

                col_f4, col_f5, col_f6 = st.columns(3)
                filtro_razao = col_f4.text_input("Razão Social")
                filtro_contraparte = col_f5.text_input("Contraparte")
                filtro_boleta = col_f6.text_input("Boleta")

                if filtro_operacao: base_exibicao = base_exibicao[base_exibicao["Operação"].isin(filtro_operacao)]
                if filtro_status: base_exibicao = base_exibicao[base_exibicao["Contrato CliqCCEE"].astype(str).isin(filtro_status)]
                if filtro_submercado: base_exibicao = base_exibicao[base_exibicao["Submercado"].astype(str).isin(filtro_submercado)]
                if filtro_razao: base_exibicao = base_exibicao[base_exibicao["Razão Social"].astype(str).str.contains(filtro_razao, case=False, na=False)]
                if filtro_contraparte: base_exibicao = base_exibicao[base_exibicao["Contraparte"].astype(str).str.contains(filtro_contraparte, case=False, na=False)]
                if filtro_boleta: base_exibicao = base_exibicao[base_exibicao["BOLETA"].astype(str).str.contains(filtro_boleta, case=False, na=False)]

                if flag_ocultar_zerados: base_exibicao = base_exibicao[base_exibicao["Volume (MWh)"] != 0.0]

                base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
                base_exibicao["Volume MWm"]   = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

                def aplicar_estilos(row):
                    rz, p, c, s = str(row.get("Razão Social", "")).strip(), str(row.get("Parte", "")).strip(), str(row.get("Contraparte", "")).strip(), str(row.get("Submercado", "")).strip()
                    if (rz, p, c, s) in st.session_state.nets_aceitos and st.session_state.nets_aceitos[(rz, p, c, s)]:
                        return ["background-color: #9370DB"] * len(row)
                    if flag_mesmo_titular and str(row.get("Razão Social", "")).strip().upper() == str(row.get("Contraparte Razão Social", "")).strip().upper():
                        return ["background-color: #FFD700"] * len(row)
                    return [""] * len(row)

                st.dataframe(base_exibicao.style.apply(aplicar_estilos, axis=1), use_container_width=True, hide_index=True)

        elif pagina == "Encontro Energético":
            pass

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)
