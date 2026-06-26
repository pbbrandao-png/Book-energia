# APP_BOOK_ENERGIA_V21 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro
# V21: + Remoção total de rateios (Auto-referência)

import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO

# Configura o limite do Pandas Styler para evitar o erro de estouro de células devido ao aumento de colunas
pd.set_option("styler.render.max_elements", 2000000)

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
    """Lê bytes de um CSV CCEE e limpa colunas."""
    df = pd.read_csv(BytesIO(bytes_csv), sep='\t', encoding='latin1', skiprows=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # Filtrar rascunhos logo na leitura economiza muita memória e processamento
    if 'SITUACAO_CONTRATO' in df.columns:
        df = df[df['SITUACAO_CONTRATO'].str.strip().str.lower() != 'rascunho']
        
    for col in ['CODIGO_CONTRATO', 'SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA', 'MWmedio', 'LIMITE_MINIMO_MODULACAO_MW', 'LIMITE_MAXIMO_MODULACAO_MW', 'TIPO_MODULACAO']:
        if col in df.columns:
            df[col] = df[col].str.strip()
            
    df['_CHAVE'] = (
        df['SIGLA_PERFIL_VENDEDOR'].fillna('')
        + df['SIGLA_PERFIL_COMPRADOR'].fillna('')
        + df['SUBMERCADO_ENTREGA'].fillna('')
    )
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
    """Mapeia os códigos da CCEE em dicionários para busca em tempo de execução O(1)."""
    if df_ccee.empty:
        return {}, {}, {}, {}, {}, {}, {}, {}
    
    # Remove duplicados mantendo o primeiro registro válido
    df_limpo = df_ccee.drop_duplicates(subset=['CODIGO_CONTRATO'])
    
    dict_chave = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo['_CHAVE']))
    dict_vend = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_VENDEDOR', '')))
    dict_comp = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_COMPRADOR', '')))
    dict_sub = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SUBMERCADO_ENTREGA', '')))
    dict_lim_min = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MINIMO_MODULACAO_MW', '-')))
    dict_lim_max = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('LIMITE_MAXIMO_MODULACAO_MW', '-')))
    dict_tipo_mod = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('TIPO_MODULACAO', '-')))
    
    # Conjunto para checar existência imediata
    set_existentes = set(df_limpo['CODIGO_CONTRATO'])
    
    return dict_chave, dict_vend, dict_comp, dict_sub, set_existentes, dict_lim_min, dict_lim_max, dict_tipo_mod


def highlight_mesmo_titular(row):
    parte = str(row.get("Parte", "")).strip().upper()
    contraparte_rs = str(row.get("Contraparte Razão Social", "")).strip().upper()
    if parte and contraparte_rs and parte == contraparte_rs:
        return ["background-color: #FFD700"] * len(row)
    return [""] * len(row)


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

        # ── EXCLUSÃO TOTAL DOS RATEIOS (PRÓPRIA REFERÊNCIA / INTRA-PORTFÓLIO) ──
        if "Parte_razao_social" in df.columns and "Contraparte_razao_social" in df.columns:
            mask_rateio_interno = df["Parte_razao_social"].astype(str).str.strip().str.upper() == df["Contraparte_razao_social"].astype(str).str.strip().str.upper()
            df = df[~mask_rateio_interno].reset_index(drop=True)

        # ── EXCLUSÃO DE RATEIOS COM Codigo_WBC == Numero_referencia_contrato E Rateio == "SIM" ──
        if "Codigo_WBC" in df.columns and "Numero_referencia_contrato" in df.columns and "Rateio" in df.columns:
            mask_rateio_duplicado = (df["Codigo_WBC"].astype(str).str.strip() == df["Numero_referencia_contrato"].astype(str).str.strip()) & (df["Rateio"].astype(str).str.strip().str.upper() == "SIM")
            df = df[~mask_rateio_duplicado].reset_index(drop=True)

        horas_mes = {
            1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720,
            7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744
        }

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
            mapa_mes_anterior = dict(zip(df_mes_anterior["BOLETA"], df_mes_anterior["Codigo_CCEE"]))
        else:
            mapa_mes_anterior = {}

        # Extração e Combinação super rápida
        csvs_matrix = extrair_csvs_zip(zip_matrix)
        csvs_bismut = extrair_csvs_zip(zip_bismut)

        df_ccee_matrix = combiner_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        df_ccee_bismut = combiner_dfs([csvs_bismut['cceal']])
        df_ccee_acr = combiner_dfs([csvs_matrix['ccear_q']])

        # CRIAÇÃO DOS ÍNDICES DE AGILIDADE
        idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext, idx_m_min, idx_m_max, idx_m_tipo = criar_indices_busca(df_ccee_matrix)
        idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext, idx_b_min, idx_b_max, idx_b_tipo = criar_indices_busca(df_ccee_bismut)
        idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext, idx_a_min, idx_a_max, idx_a_tipo = criar_indices_busca(df_ccee_acr)

        mapa_energia = {
            "Incentivada 50%": "Incentivada-I5", "Cogeração Qualificada 50%": "Incentivada-CQ5",
            "Incentivada 100%": "Incentivada-I1", "Convencional": "Convencional", "Incentivada 0%": "Incentivada-I0"
        }
        mapa_submercado = {"Sul": "SUL", "S": "SUL", "SE/CO": "SUDESTE", "N": "NORTE", "NE": "NORDESTE"}
        mapa_modulacao = {"F - Flat": "FLAT", "C - Carga": "CARGA", "DECLARADO": "DECLARADA", "G - Geração": "GERAÇÃO"}

        df["Suprimento_inicio"] = pd.to_datetime(df["Suprimento_inicio"], errors="coerce")
        df["Suprimento_termino"] = pd.to_datetime(df["Suprimento_termino"], errors="coerce")

        dias_periodo = (df["Suprimento_termino"] - df["Suprimento_inicio"]).dt.days + 1
        cp_lp = dias_periodo.apply(lambda x: "CP" if x <= 31 else "LP")
        horas_por_linha = df["Mes"].map(horas_mes)
        volume_mwm = (df["QuantAtualizada"] / horas_por_linha).round(6)

        # Converter Codigo_CCEE para string antes de usar
        codigo_ccee_str = df["Codigo_CCEE"].fillna("").astype(str).str.strip()
        codigo_ccee_str = codigo_ccee_str.replace("nan", "-")
        codigo_ccee_str = codigo_ccee_str.replace("", "-")
        codigo_ccee_str = codigo_ccee_str.apply(lambda x: "-" if x == "0" or x == "" else x)

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
        base["CliqCCEE Paradigma"]             = codigo_ccee_str
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["% Modulação Mínima"]             = df["FlexLimite_modulacaoMin"].fillna("-")
        base["% Modulação Máxima"]             = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-").astype(str)
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str)
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str)
        base["Contrato CliqCCEE"]              = "-"

        csvs_disponiveis = any([not df_ccee_matrix.empty, not df_ccee_bismut.empty, not df_ccee_acr.empty])

        if csvs_disponiveis:
            BISMUT_NOME_UPPER = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."
            
            def calcular_contrato_cliqccee_fast(row):
                try:
                    b_int = int(float(str(row["BOLETA"]).strip()))
                except:
                    b_int = -1
                
                if b_int in BOLETAS_ACR:
                    d_ch, s_ext = idx_a_chave, set_a_ext
                elif str(row["Parte"]).strip().upper() == BISMUT_NOME_UPPER:
                    d_ch, s_ext = idx_b_chave, set_b_ext
                else:
                    d_ch, s_ext = idx_m_chave, set_m_ext
                
                chave_esp = str(row["Vendedor"]).strip() + str(row["Comprador"]).strip() + str(row["Submercado"]).strip()
                
                c_ant = str(row["Contrato CliqCCEE mês anterior"]).strip()
                if c_ant in s_ext:
                    return c_ant if d_ch.get(c_ant) == chave_esp else 'Verificar'
                
                c_par = str(row["CliqCCEE Paradigma"]).strip()
                if c_par in s_ext:
                    return c_par if d_ch.get(c_par) == chave_esp else 'Verificar'
                
                return '-'

            base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee_fast, axis=1).astype(str)

        # ── COLUNA: Volume Book ──────────────────────────────────────────────────
        _vol_mwm_num = pd.to_numeric(base["Volume MWm"], errors="coerce")
        _mask_valido_book = _vol_mwm_num.notna() & (base["Volume MWm"].astype(str).str.strip() != "-")
        _df_book = base[["Contrato CliqCCEE"]].copy()
        _df_book["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
        _soma_book = _df_book.groupby("Contrato CliqCCEE")["_vol_num"].transform("sum")
        base["Volume Book"] = _soma_book

        # ── CÁLCULO DAS COLUNAS DE MODULAÇÃO BOOK E CCEE ──────────────────
        _vol_book_num = pd.to_numeric(base["Volume Book"], errors="coerce").fillna(0.0)
        _num_mod_min = pd.to_numeric(base["% Modulação Mínima"], errors="coerce").fillna(0.0)
        _num_mod_max = pd.to_numeric(base["% Modulação Máxima"], errors="coerce").fillna(0.0)

        _mask_tem_contrato = ~base["Contrato CliqCCEE"].astype(str).str.strip().isin(["", "-", "None", "nan", "Verificar"])
        _mask_calcular_min = _mask_tem_contrato & (_num_mod_min > 0.0)
        _mask_calcular_max = _mask_tem_contrato & (_num_mod_max > 0.0)

        base["Modulação Mínima"] = (_vol_book_num * (1 - (_num_mod_min / 100))).where(_mask_calcular_min, "-")
        base["Modulação Máxima"] = (_vol_book_num * (1 + (_num_mod_max / 100))).where(_mask_calcular_max, "-")

        if csvs_disponiveis:
            def buscar_campo_ccee(row, dict_m, dict_b, dict_a):
                cod = str(row["Contrato CliqCCEE"]).strip()
                if cod in ["", "-", "None", "nan", "Verificar"]:
                    return "-"
                try:
                    b_int = int(float(str(row["BOLETA"]).strip()))
                except:
                    b_int = -1
                
                if b_int in BOLETAS_ACR:
                    d_field = dict_a
                elif str(row["Parte"]).strip().upper() == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.":
                    d_field = dict_b
                else:
                    d_field = dict_m
                
                return d_field.get(cod, "-")

            base["Modulação Mínima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_min, idx_b_min, idx_a_min), axis=1)
            base["Modulação Máxima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_max, idx_b_max, idx_a_max), axis=1)
            base["Modulação CCEE"]        = base.apply(lambda r: buscar_campo_ccee(r, idx_m_tipo, idx_b_tipo, idx_a_tipo), axis=1)
        else:
            base["Modulação Mínima CCEE"] = "-"
            base["Modulação Máxima CCEE"] = "-"
            base["Modulação CCEE"]        = "-"

        _tol_mod = 1e-4
        
        # Check Modulação Mínima
        base["Check Modulação Mínima"] = "-"
        _mod_min_cc_str = base.loc[_mask_calcular_min, "Modulação Mínima CCEE"].astype(str).str.replace(",", ".", regex=False)
        _mod_min_cc = pd.to_numeric(_mod_min_cc_str, errors="coerce")
        
        _mask_min_valid = _mask_calcular_min & _mod_min_cc.notna()
        if _mask_min_valid.any():
            _diff_min = pd.to_numeric(base.loc[_mask_min_valid, "Modulação Mínima"]) - _mod_min_cc.loc[_mask_min_valid]
            base.loc[_mask_min_valid, "Check Modulação Mínima"] = "OK"
            base.loc[_mask_min_valid & (_diff_min > _tol_mod), "Check Modulação Mínima"] = "Book maior"
            base.loc[_mask_min_valid & (_diff_min < -_tol_mod), "Check Modulação Mínima"] = "CCEE maior"

        # Check Modulação Máxima
        base["Check Modulação Máxima"] = "-"
        _mod_max_cc_str = base.loc[_mask_calcular_max, "Modulação Máxima CCEE"].astype(str).str.replace(",", ".", regex=False)
        _mod_max_cc = pd.to_numeric(_mod_max_cc_str, errors="coerce")
        
        _mask_max_valid = _mask_calcular_max & _mod_max_cc.notna()
        if _mask_max_valid.any():
            _diff_max = pd.to_numeric(base.loc[_mask_max_valid, "Modulação Máxima"]) - _mod_max_cc.loc[_mask_max_valid]
            base.loc[_mask_max_valid, "Check Modulação Máxima"] = "OK"
            base.loc[_mask_max_valid & (_diff_max > _tol_mod), "Check Modulação Máxima"] = "Book maior"
            base.loc[_mask_max_valid & (_diff_max < -_tol_mod), "Check Modulação Máxima"] = "CCEE maior"

        # Check Modulação Tipo
        base["Check Modulação"] = "-"
        _mask_tipo_valid = _mask_tem_contrato & (~base["Modulação CCEE"].astype(str).str.strip().isin(["", "-", "None", "nan"]))
        if _mask_tipo_valid.any():
            _mask_div_tipo = base["Modulação WBC"].astype(str).str.strip().str.upper() != base["Modulação CCEE"].astype(str).str.strip().str.upper()
            base.loc[_mask_tipo_valid, "Check Modulação"] = "OK"
            base.loc[_mask_tipo_valid & _mask_div_tipo, "Check Modulação"] = "Divergente"

        _ordem_colunas = [
            "BOLETA", "Operação", "Tipo de Energia", "Parte", "Contraparte Razão Social", "Contraparte",
            "CP/LP", "CNPJ CONTRAPARTE", "Submercado", "Volume (MWh)", "Volume MWm", "CliqCCEE Paradigma",
            "Modulação WBC", "% Modulação Mínima", "Modulação Mínima", "Modulação Mínima CCEE", "Check Modulação Mínima",
            "% Modulação Máxima", "Modulação Máxima", "Modulação Máxima CCEE", "Check Modulação Máxima",
            "Modulação CCEE", "Check Modulação",
            "Contrato CliqCCEE mês anterior", "Vendedor", "Comprador", "Contrato CliqCCEE",
            "Volume Book", "Volume CCEE", "Check Volume", "Volume Global", "Volume Global CCEE", "Check Volume Global"
        ]
        base = base[[c for c in _ordem_colunas if c in base.columns]]

        # ── COLUNA: Volume CCEE ──────────────────────────────────────────────────
        if csvs_disponiveis:
            _lista_dfs_ccee_vol = []
            for _df_src in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
                if _df_src is not None and not _df_src.empty and "CODIGO_CONTRATO" in _df_src.columns and "MWmedio" in _df_src.columns:
                    _tmp = _df_src[["CODIGO_CONTRATO", "MWmedio"]].copy()
                    _tmp["MWmedio"] = _tmp["MWmedio"].astype(str).str.strip().str.replace(",", ".", regex=False)
                    _tmp["MWmedio"] = pd.to_numeric(_tmp["MWmedio"], errors="coerce").fillna(0.0)
                    _lista_dfs_ccee_vol.append(_tmp)
            if _lista_dfs_ccee_vol:
                _df_ccee_vol = pd.concat(_lista_dfs_ccee_vol, ignore_index=True)
                _vol_ccee_por_contrato = _df_ccee_vol.groupby("CODIGO_CONTRATO")["MWmedio"].sum()
                base["Volume CCEE"] = base["Contrato CliqCCEE"].map(_vol_ccee_por_contrato).fillna(0.0)
            else:
                base["Volume CCEE"] = 0.0
        else:
            base["Volume CCEE"] = 0.0

        # ── COLUNA: Check Volume ─────────────────────────────────────────────────
        _tol = 1e-6
        _vb = pd.to_numeric(base["Volume Book"], errors="coerce").fillna(0.0)
        _vc = pd.to_numeric(base["Volume CCEE"], errors="coerce").fillna(0.0)
        _diff_vol = _vb - _vc
        base["Check Volume"] = "OK"
        base.loc[_diff_vol > _tol, "Check Volume"] = "Book maior"
        base.loc[_diff_vol < -_tol, "Check Volume"] = "CCEE maior"

        # ── COLUNA: Volume Global ────────────────────────────────────────────────
        _df_global = base[["Vendedor", "Comprador", "Submercado"]].copy()
        _df_global["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
        _soma_global = _df_global.groupby(["Vendedor", "Comprador", "Submercado"])["_vol_num"].transform("sum")
        base["Volume Global"] = _soma_global

        # ── COLUNA: Volume Global CCEE ───────────────────────────────────────────
        if csvs_disponiveis:
            _lista_dfs_global_ccee = []
            for _df_src in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
                if _df_src is not None and not _df_src.empty and "MWmedio" in _df_src.columns:
                    _cols_need = ["SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR", "SUBMERCADO_ENTREGA", "MWmedio"]
                    if all(c in _df_src.columns for c in _cols_need):
                        _tmp2 = _df_src[_cols_need].copy()
                        _tmp2["MWmedio"] = _tmp2["MWmedio"].astype(str).str.strip().str.replace(",", ".", regex=False)
                        _tmp2["MWmedio"] = pd.to_numeric(_tmp2["MWmedio"], errors="coerce").fillna(0.0)
                        _lista_dfs_global_ccee.append(_tmp2)
            if _lista_dfs_global_ccee:
                _df_gc = pd.concat(_lista_dfs_global_ccee, ignore_index=True)
                _gc_sum = _df_gc.groupby(["SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR", "SUBMERCADO_ENTREGA"])["MWmedio"].sum()
                _gc_sum.index.names = ["Vendedor", "Comprador", "Submercado"]
                _gc_sum = _gc_sum.reset_index()
                base = base.merge(_gc_sum, on=["Vendedor", "Comprador", "Submercado"], how="left")
                base.rename(columns={"MWmedio": "Volume Global CCEE"}, inplace=True)
                base["Volume Global CCEE"] = base["Volume Global CCEE"].fillna(0.0)
            else:
                base["Volume Global CCEE"] = 0.0
        else:
            base["Volume Global CCEE"] = 0.0

        # ── COLUNA: Check Volume Global ──────────────────────────────────────────
        _vg = pd.to_numeric(base["Volume Global"], errors="coerce").fillna(0.0)
        _vgc = pd.to_numeric(base["Volume Global CCEE"], errors="coerce").fillna(0.0)
        _diff_global = _vg - _vgc
        base["Check Volume Global"] = "OK"
        base.loc[_diff_global > _tol, "Check Volume Global"] = "Book maior"
        base.loc[_diff_global < -_tol, "Check Volume Global"] = "CCEE maior"

        compras_net = base[base["Operação"] == "Compra"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"].sum().rename(columns={"Volume (MWh)": "Compra (MWh)"})
        vendas_net = base[base["Operação"] == "Venda"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"].sum().rename(columns={"Volume (MWh)": "Venda (MWh)"})
        nets = compras_net.merge(vendas_net, on=["Parte", "Contraparte", "Submercado", "Tipo de Energia"], how="inner")

        if pagina == "Base Conferência":
            st.subheader("Base Conferência")
            total_contratos = len(base)
            total_compras = len(base[base['Operação'].str.upper() == 'COMPRA'])
            total_vendas = len(base[base['Operação'].str.upper() == 'VENDA'])

            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric(label="Total de Contratos (Sem Rateios)", value=total_contratos)
            col_metric2.metric(label="Contratos de Compra 📥", value=total_compras)
            col_metric3.metric(label="Contratos de Venda 📤", value=total_vendas)
            st.markdown("---")

            # ── RESUMO DE NETs ────────────────────────────────────────────────────
            # Reutiliza integralmente a lógica de formação dos NETs já existente (inner merge de compras x vendas)
            mes_referencia_net = int(df["Mes"].dropna().iloc[0]) if "Mes" in df.columns and not df["Mes"].dropna().empty else 1
            horas_net = horas_mes.get(mes_referencia_net, 744)

            compras_net_mwm = (
                base[base["Operação"] == "Compra"]
                .groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume MWm"]
                .sum()
                .rename(columns={"Volume MWm": "Compra (MWm)"})
            )
            vendas_net_mwm = (
                base[base["Operação"] == "Venda"]
                .groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume MWm"]
                .sum()
                .rename(columns={"Volume MWm": "Venda (MWm)"})
            )
            nets_resumo = compras_net_mwm.merge(
                vendas_net_mwm,
                on=["Parte", "Contraparte", "Submercado", "Tipo de Energia"],
                how="inner"
            )

            if not nets_resumo.empty:
                # Volume NET = mínimo entre compra e venda (volume compensável)
                nets_resumo["Volume NET (MWm)"] = nets_resumo[["Compra (MWm)", "Venda (MWm)"]].min(axis=1)

                # Responsável pelo Ajuste: quem tem o maior volume deve ajustar o excedente no Cliq
                def responsavel_ajuste(row):
                    saldo = row["Compra (MWm)"] - row["Venda (MWm)"]
                    if abs(saldo) < 1e-9:
                        return "ZERADO"
                    elif saldo > 0:
                        return row["Contraparte"]
                    else:
                        return row["Parte"]

                nets_resumo["Responsável pelo Ajuste"] = nets_resumo.apply(responsavel_ajuste, axis=1)

                # Vendedor e Comprador (perspectiva do NET: quem vende para quem)
                # Parte é sempre quem aparece nos dois lados; Contraparte é o outro lado
                nets_resumo["Vendedor"] = nets_resumo.apply(
                    lambda r: r["Parte"] if r["Venda (MWm)"] >= r["Compra (MWm)"] else r["Contraparte"], axis=1
                )
                nets_resumo["Comprador"] = nets_resumo.apply(
                    lambda r: r["Parte"] if r["Compra (MWm)"] >= r["Venda (MWm)"] else r["Contraparte"], axis=1
                )

                # Compra Cliq e Venda Cliq por NET específico (usando CSVs já carregados)
                def calcular_cliq_net(row):
                    parte = str(row["Parte"]).strip()
                    contraparte = str(row["Contraparte"]).strip()
                    submercado = str(row["Submercado"]).strip()

                    compra_cliq = 0.0
                    venda_cliq = 0.0

                    dfs_ccee = []
                    if csvs_disponiveis:
                        for _df_src in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
                            if _df_src is not None and not _df_src.empty:
                                dfs_ccee.append(_df_src)

                    for _df_src in dfs_ccee:
                        cols_need = ["SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR", "SUBMERCADO_ENTREGA", "MWmedio"]
                        if not all(c in _df_src.columns for c in cols_need):
                            continue
                        _tmp = _df_src[cols_need].copy()
                        _tmp["MWmedio"] = _tmp["MWmedio"].astype(str).str.strip().str.replace(",", ".", regex=False)
                        _tmp["MWmedio"] = pd.to_numeric(_tmp["MWmedio"], errors="coerce").fillna(0.0)

                        # Compra Cliq: Parte como comprador, contraparte como vendedor
                        mask_compra = (
                            (_tmp["SIGLA_PERFIL_COMPRADOR"].str.strip() == parte) &
                            (_tmp["SIGLA_PERFIL_VENDEDOR"].str.strip() == contraparte) &
                            (_tmp["SUBMERCADO_ENTREGA"].str.strip() == submercado)
                        )
                        compra_cliq += _tmp.loc[mask_compra, "MWmedio"].sum()

                        # Venda Cliq: Parte como vendedor, contraparte como comprador
                        mask_venda = (
                            (_tmp["SIGLA_PERFIL_VENDEDOR"].str.strip() == parte) &
                            (_tmp["SIGLA_PERFIL_COMPRADOR"].str.strip() == contraparte) &
                            (_tmp["SUBMERCADO_ENTREGA"].str.strip() == submercado)
                        )
                        venda_cliq += _tmp.loc[mask_venda, "MWmedio"].sum()

                    return pd.Series({"Compra Cliq (MWm)": compra_cliq, "Venda Cliq (MWm)": venda_cliq})

                cliq_cols = nets_resumo.apply(calcular_cliq_net, axis=1)
                nets_resumo["Compra Cliq (MWm)"] = cliq_cols["Compra Cliq (MWm)"]
                nets_resumo["Venda Cliq (MWm)"] = cliq_cols["Venda Cliq (MWm)"]

                # Inicializar estado de checkboxes no session_state
                if "net_efetivados" not in st.session_state:
                    st.session_state["net_efetivados"] = {}

                net_keys = list(nets_resumo.apply(
                    lambda r: f"{r['Parte']}|{r['Contraparte']}|{r['Submercado']}|{r['Tipo de Energia']}", axis=1
                ))
                for k in net_keys:
                    if k not in st.session_state["net_efetivados"]:
                        st.session_state["net_efetivados"][k] = False

                efetivados_flag = [st.session_state["net_efetivados"].get(k, False) for k in net_keys]
                nets_resumo["_key"] = net_keys
                nets_resumo["_efetivado"] = efetivados_flag

                # Status
                def calcular_status(row):
                    if not row["_efetivado"]:
                        return "Não efetivado"
                    vol_net = row["Volume NET (MWm)"]
                    compra_cliq = row["Compra Cliq (MWm)"]
                    venda_cliq = row["Venda Cliq (MWm)"]
                    tol = 1e-6

                    if abs(compra_cliq - venda_cliq) > tol:
                        return "🔴 Divergência entre Compra Cliq e Venda Cliq"

                    ajuste = (compra_cliq + venda_cliq) / 2 if abs(compra_cliq - venda_cliq) < tol else compra_cliq

                    if abs(ajuste) < tol and abs(vol_net) > tol:
                        return "🟡 Aguardando ajuste"
                    if abs(ajuste - vol_net) < tol:
                        return "✅ OK"
                    if ajuste > vol_net + tol:
                        return "🔴 Volume ajustado maior que o esperado"
                    if ajuste < vol_net - tol:
                        if ajuste < tol:
                            return "🟡 Aguardando ajuste"
                        return "🟠 Ajuste parcial"
                    return "✅ OK"

                nets_resumo["Status"] = nets_resumo.apply(calcular_status, axis=1)

                with st.expander("📋 Resumo de NETs", expanded=False):
                    st.caption(f"{len(nets_resumo)} NET(s) identificado(s)")

                    for idx, row in nets_resumo.iterrows():
                        key = row["_key"]
                        col_chk, col_info = st.columns([1, 11])
                        with col_chk:
                            novo_val = st.checkbox(
                                "Efetivado",
                                value=st.session_state["net_efetivados"].get(key, False),
                                key=f"net_chk_{key}",
                                label_visibility="collapsed"
                            )
                            if novo_val != st.session_state["net_efetivados"].get(key, False):
                                st.session_state["net_efetivados"][key] = novo_val
                                st.rerun()

                    # Montar tabela de exibição
                    display_nets = nets_resumo[[
                        "_efetivado", "Parte", "Vendedor", "Comprador",
                        "Compra (MWm)", "Venda (MWm)", "Volume NET (MWm)",
                        "Responsável pelo Ajuste", "Compra Cliq (MWm)", "Venda Cliq (MWm)", "Status"
                    ]].copy()
                    display_nets.rename(columns={"_efetivado": "Efetivado"}, inplace=True)

                    for col_mwm in ["Compra (MWm)", "Venda (MWm)", "Volume NET (MWm)", "Compra Cliq (MWm)", "Venda Cliq (MWm)"]:
                        display_nets[col_mwm] = display_nets[col_mwm].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

                    st.dataframe(display_nets, use_container_width=True, hide_index=True)

                # Identificar boletas pertencentes a NETs efetivados para highlight roxo
                nets_efetivados_keys = {k for k, v in st.session_state["net_efetivados"].items() if v}

            else:
                nets_efetivados_keys = set()
                with st.expander("📋 Resumo de NETs", expanded=False):
                    st.info("Nenhuma possibilidade de NET identificada.")

            st.markdown("---")

            col_flag1, col_flag2 = st.columns(2)
            with col_flag1: flag_mesmo_titular = st.toggle("🟡 Ocultar IntraPortifólio Visualmente", value=True)
            with col_flag2: flag_ocultar_zerados = st.toggle("🚫 Ocultar contratos zerados (Volume MWh = 0)", value=False)

            base_exibicao = base.copy()
            st.markdown("### 🔎 Filtros")
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1: filtro_operacao = st.multiselect("Operação", options=sorted(base_exibicao["Operação"].dropna().unique()), default=[])
            with col_f2: filtro_status = st.multiselect("Contrato CliqCCEE", options=sorted(base_exibicao["Contrato CliqCCEE"].dropna().astype(str).unique()), default=[])
            with col_f3: filtro_submercado = st.multiselect("Submercado", options=sorted(base_exibicao["Submercado"].dropna().astype(str).unique()), default=[])

            col_f4, col_f5, col_f6 = st.columns(3)
            with col_f4: filtro_parte = st.text_input("Parte")
            with col_f5: filtro_contraparte = st.text_input("Contraparte")
            with col_f6: filtro_boleta = st.text_input("Boleta")

            if filtro_operacao: base_exibicao = base_exibicao[base_exibicao["Operação"].isin(filtro_operacao)]
            if filtro_status: base_exibicao = base_exibicao[base_exibicao["Contrato CliqCCEE"].astype(str).isin(filtro_status)]
            if filtro_submercado: base_exibicao = base_exibicao[base_exibicao["Submercado"].astype(str).isin(filtro_submercado)]
            if filtro_parte: base_exibicao = base_exibicao[base_exibicao["Parte"].astype(str).str.contains(filtro_parte, case=False, na=False)]
            if filtro_contraparte: base_exibicao = base_exibicao[base_exibicao["Contraparte"].astype(str).str.contains(filtro_contraparte, case=False, na=False)]
            if filtro_boleta: base_exibicao = base_exibicao[base_exibicao["BOLETA"].astype(str).str.contains(filtro_boleta, case=False, na=False)]

            if flag_ocultar_zerados: base_exibicao = base_exibicao[base_exibicao["Volume (MWh)"] != 0.0]

            base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            base_exibicao["Volume MWm"]   = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            
            for c_format in ["Modulação Mínima", "Modulação Máxima"]:
                if c_format in base_exibicao.columns:
                    base_exibicao[c_format] = base_exibicao[c_format].map(lambda x: f"{x:.4f}" if isinstance(x, (int, float)) else x)

            st.caption(f"{len(base_exibicao):,} registros encontrados")

            # Converter para string apenas para exibição
            colunas_texto = [
                "BOLETA", "Operação", "Tipo de Energia", "Parte", "Contraparte Razão Social",
                "Contraparte", "CP/LP", "CNPJ CONTRAPARTE", "Submercado", "CliqCCEE Paradigma",
                "Contrato CliqCCEE mês anterior", "Contrato CliqCCEE", "Modulação WBC",
                "% Modulação Mínima", "% Modulação Máxima", "Modulação Mínima", "Modulação Máxima",
                "Modulação Mínima CCEE", "Modulação Máxima CCEE", "Check Modulação Mínima",
                "Check Modulação Máxima", "Modulação CCEE", "Check Modulação", "Vendedor", "Comprador"
            ]
            for col in colunas_texto:
                if col in base_exibicao.columns:
                    base_exibicao[col] = base_exibicao[col].astype(str)

            # Highlight roxo para contratos pertencentes a NETs efetivados
            def highlight_net_efetivado(row):
                parte = str(row.get("Parte", "")).strip()
                contraparte = str(row.get("Contraparte", "")).strip()
                submercado = str(row.get("Submercado", "")).strip()
                tipo_energia = str(row.get("Tipo de Energia", "")).strip()
                key = f"{parte}|{contraparte}|{submercado}|{tipo_energia}"
                if key in nets_efetivados_keys:
                    return ["background-color: #7B2D8B; color: white"] * len(row)
                return [""] * len(row)

            if flag_mesmo_titular and nets_efetivados_keys:
                def highlight_combinado(row):
                    parte = str(row.get("Parte", "")).strip()
                    contraparte = str(row.get("Contraparte", "")).strip()
                    submercado = str(row.get("Submercado", "")).strip()
                    tipo_energia = str(row.get("Tipo de Energia", "")).strip()
                    key = f"{parte}|{contraparte}|{submercado}|{tipo_energia}"
                    if key in nets_efetivados_keys:
                        return ["background-color: #7B2D8B; color: white"] * len(row)
                    contraparte_rs = str(row.get("Contraparte Razão Social", "")).strip().upper()
                    if parte and contraparte_rs and parte.upper() == contraparte_rs:
                        return ["background-color: #FFD700"] * len(row)
                    return [""] * len(row)
                styled = base_exibicao.style.apply(highlight_combinado, axis=1)
                st.dataframe(styled, use_container_width=True, hide_index=True)
            elif flag_mesmo_titular:
                styled = base_exibicao.style.apply(highlight_mesmo_titular, axis=1)
                st.dataframe(styled, use_container_width=True, hide_index=True)
            elif nets_efetivados_keys:
                styled = base_exibicao.style.apply(highlight_net_efetivado, axis=1)
                st.dataframe(styled, use_container_width=True, hide_index=True)
            else:
                st.dataframe(base_exibicao, use_container_width=True, hide_index=True)

            base_download = base.copy()
            if flag_ocultar_zerados: base_download = base_download[base_download["Volume (MWh)"] != 0.0]
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer: base_download.to_excel(writer, sheet_name="Base Conferência", index=False)
            st.download_button("📥 Download Base Conferência", data=output.getvalue(), file_name="Base_Conferencia.xlsx")

            if csvs_disponiveis:
                st.markdown("---")
                lista_divergencias = []
                lista_sem_match_nenhum = []

                for _, row in base.iterrows():
                    try: b_int = int(float(str(row["BOLETA"]).strip()))
                    except: b_int = -1

                    # Ignora contratos com volume zerado
                    volume = float(row.get("Volume (MWh)", 0))
                    if volume == 0:
                        continue

                    if b_int in BOLETAS_ACR:
                        d_ch, d_v, d_c, d_s, s_ext = idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext
                    elif str(row["Parte"]).strip().upper() == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.":
                        d_ch, d_v, d_c, d_s, s_ext = idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext
                    else:
                        d_ch, d_v, d_c, d_s, s_ext = idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext

                    v_b, c_b, s_b = str(row["Vendedor"]).strip(), str(row["Comprador"]).strip(), str(row["Submercado"]).strip()
                    chave_esp = v_b + c_b + s_b

                    cods = [str(row.get(c, "")).strip() for c in ["Contrato CliqCCEE", "Contrato CliqCCEE mês anterior", "CliqCCEE Paradigma"]]
                    cods_validos = [c for c in cods if c not in ('', '-', 'None', 'nan', 'Verificar')]

                    cod_encontrado = None
                    for c in cods_validos:
                        if c in s_ext:
                            cod_encontrado = c
                            break

                    if not cod_encontrado:
                        if chave_esp in d_ch.values():
                            status, justificativa = "OK", None
                        else:
                            status, justificativa = "SEM_MATCH", "Contrato inexistente no CSV CCEE"
                    else:
                        v_c, c_c, s_c = d_v.get(cod_encontrado, ''), d_c.get(cod_encontrado, ''), d_s.get(cod_encontrado, '')

                        divs = []
                        if v_b != v_c: divs.append("Vendedor")
                        if c_b != c_c: divs.append("Comprador")
                        if s_b != s_c: divs.append(f"Submercado (Boleta={s_b} | CSV={s_c})")

                        if not divs:
                            status, justificativa = "OK", None
                        else:
                            status = "ERRO"
                            if len(divs) == 1: justificativa = f"Divergência de {divs[0]}"
                            elif len(divs) == 2: justificativa = f"Divergência de {divs[0]} e {divs[1]}"
                            else: justificativa = f"Divergência de {divs[0]}, {divs[1]} e {divs[2]}"

                    if status in ("ERRO", "SEM_MATCH"):
                        item = {"Boleta": row["BOLETA"], "Vendedor": row["Vendedor"], "Comprador": row["Comprador"], "Mensagem": justificativa}
                        if status == "ERRO": lista_divergencias.append(item)
                        else: lista_sem_match_nenhum.append(item)

                df_divergencias = pd.DataFrame(lista_divergencias, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])
                df_sem_match_nenhum = pd.DataFrame(lista_sem_match_nenhum, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])

                st.subheader("❌ Contratos com Divergência (Existem no CSV, mas dados não batem)")
                st.dataframe(df_divergencias, use_container_width=True, hide_index=True)

                output_div = BytesIO()
                with pd.ExcelWriter(output_div, engine="openpyxl") as writer: df_divergencias.to_excel(writer, sheet_name="Divergencias", index=False)
                st.download_button("📥 Download Contratos com Divergência", data=output_div.getvalue(), file_name="Contratos_com_Divergencia.xlsx")
                st.markdown("---")

                st.subheader("🔍 Contratos Sem Match Nenhum (Inexistentes no CSV CCEE)")
                st.dataframe(df_sem_match_nenhum, use_container_width=True, hide_index=True)

                output_sm = BytesIO()
                with pd.ExcelWriter(output_sm, engine="openpyxl") as writer: df_sem_match_nenhum.to_excel(writer, sheet_name="Sem Match Nenhum", index=False)
                st.download_button("📥 Download Contratos Sem Match Nenhum", data=output_sm.getvalue(), file_name="Contratos_Sem_Match_Nenhum.xlsx")

        elif pagina == "Encontro Energético":
            st.subheader("🤝 Encontro Energético")
            parte = st.selectbox("Parte", sorted(nets["Parte"].dropna().unique()))
            df_parte = nets[nets["Parte"] == parte]
            contraparte = st.selectbox("Contraparte", sorted(df_parte["Contraparte"].dropna().unique()))
            df_contraparte = df_parte[df_parte["Contraparte"] == contraparte]
            submercado = st.selectbox("Submercado", sorted(df_contraparte["Submercado"].dropna().unique()))
            df_sub = df_contraparte[df_contraparte["Submercado"] == submercado]
            tipo_energia = st.selectbox("Tipo de Energia", sorted(df_sub["Tipo de Energia"].dropna().unique()))

            encontro = base[(base["Parte"] == parte) & (base["Contraparte"] == contraparte) & (base["Submercado"] == submercado) & (base["Tipo de Energia"] == tipo_energia)]
            compras_calc = encontro[encontro["Operação"] == "Compra"]
            vendas_calc  = encontro[encontro["Operação"] == "Venda"]

            compras, vendas = compras_calc.copy(), vendas_calc.copy()
            compras["Volume (MWh)"] = compras["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            compras["Volume MWm"]   = compras["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            vendas["Volume (MWh)"]  = vendas["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            vendas["Volume MWm"]    = vendas["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

            st.markdown("## COMPRAS")
            st.dataframe(compras[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)
            st.markdown("## VENDAS")
            st.dataframe(vendas[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)

            total_compra, total_venda = compras_calc["Volume (MWh)"].sum(), vendas_calc["Volume (MWh)"].sum()
            saldo = total_compra - total_venda
            total_compra_mwm, total_venda_mwm = compras_calc["Volume MWm"].sum(), vendas_calc["Volume MWm"].sum()
            mes_referencia = int(df["Mes"].dropna().iloc[0])
            saldo_mwm = saldo / horas_mes.get(mes_referencia, 744)

            ajuste = contraparte if saldo > 0 else parte if saldo < 0 else "ZERADO"
            resumo = pd.DataFrame({"Tipo": ["Compras", "Vendas", "Saldo"], "MWh": [f"{total_compra:.3f}", f"{total_venda:.3f}", f"{saldo:.3f}"], "MWm": [f"{total_compra_mwm:.6f}", f"{total_venda_mwm:.6f}", f"{saldo_mwm:.6f}"]})
            st.markdown("## RESUMO")
            st.dataframe(resumo, hide_index=True, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1: st.metric("Quem Ajusta", ajuste)
            with c2: st.metric("Volume a Ajustar (MWm)", f"{abs(saldo_mwm):.6f}")

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)
