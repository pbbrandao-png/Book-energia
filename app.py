# APP_BOOK_ENERGIA_V22 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro
# V21: + Remoção total de rateios (Auto-referência)
# V22: + Identificação e Filtro de Varejistas (MATRIX VAR / BISMUT VAR) + Correção Bug Comprapor

import streamlit as st
import pandas as pd
import zipfile
import numpy as np
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


# Partes elegíveis para a regra InterCompany
_PARTES_INTERCOMPANY = {
    "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.",
    "GET COMERCIALIZADORA DE ENERGIA S.A.",
    "ARGENTUM COMERCIALIZADORA DE ENERGIA LTDA.",
}


def aplicar_zerar_intercompany(base: pd.DataFrame):
    """
    Recebe uma cópia da Base Conferência e zera Volume (MWh) e Volume MWm
    dos contratos InterCompany, conforme regra:

    - Parte belongs to _PARTES_INTERCOMPANY
    - Contraparte (sigla CCEE) starts with "MATRIX"
    - BUT DOES NOT start with "MATRIX VAR"

    Retorna (base_modificada, mask_intercompany).
    """
    base = base.copy()

    parte_upper = base["Parte"].astype(str).str.strip().str.upper()
    contra_upper = base["Contraparte"].astype(str).str.strip().str.upper()

    mask_parte     = parte_upper.isin(_PARTES_INTERCOMPANY)
    mask_matrix    = contra_upper.str.startswith("MATRIX")
    mask_matrix_var = contra_upper.str.startswith("MATRIX VAR")

    mask_intercompany = mask_parte & mask_matrix & ~mask_matrix_var

    base.loc[mask_intercompany, "Volume (MWh)"] = 0.0
    base.loc[mask_intercompany, "Volume MWm"]   = 0.0

    return base, mask_intercompany


# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Book Energia", layout="wide")

pagina = st.sidebar.radio("Menu", ["Base Conferência", "Encontro Energético", "Arquivos CCEE"])
st.sidebar.markdown("---")

st.title("📊 Book Energia")

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx", "xlsx"])
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

        # ── INCLUSÃO: GARANTIR ORDENAÇÃO CRESCENTE E REMOVER BOLETAS DUPLICADAS ──
        if "Codigo_WBC" in df.columns:
            df = df.iloc[pd.to_numeric(df["Codigo_WBC"], errors="coerce").argsort()].reset_index(drop=True)
            df = df.drop_duplicates(subset=["Codigo_WBC"], keep="first").reset_index(drop=True)

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

        # ── LOGICA PARA DEFINIR SE É VAREJISTA (MATRIX VAR OU BISMUT VAR) ──
        p_upper = base["Parte"].astype(str).str.strip().str.upper()
        c_upper = base["Contraparte"].astype(str).str.strip().str.upper()
        
        mask_varejista = (
            p_upper.str.startswith("MATRIX VAR") | 
            p_upper.str.startswith("BISMUT VAR") | 
            c_upper.str.startswith("MATRIX VAR") | 
            c_upper.str.startswith("BISMUT VAR")
        )
        base["Varejista"] = np.where(mask_varejista, "Sim", "Não")

        csvs_disponiveis = any([not df_ccee_matrix.empty, not df_ccee_bismut.empty, not df_ccee_acr.empty])

        if pagina == "Base Conferência":
            st.subheader("Base Conferência")
            total_contratos = len(base)
            total_compras = len(base[base['Operação'].str.upper() == 'COMPRA'])
            total_vendas = len(base[base['Operação'].str.upper() == 'VENDA'])

            col_metric1, col_metric2, col_metric3 = st.columns(3)
            col_metric1.metric(label="Total de Contratos (Sem Rateios/Duplicadas)", value=total_contratos)
            col_metric2.metric(label="Contratos de Compra 📥", value=total_compras)
            col_metric3.metric(label="Contratos de Venda 📤", value=total_vendas)
            st.markdown("---")

            col_flag1, col_flag2, col_flag3 = st.columns(3)
            with col_flag1: flag_mesmo_titular = st.toggle("🟡 Ocultar IntraPortifólio Visualmente", value=True)
            with col_flag2: flag_ocultar_zerados = st.toggle("🚫 Ocultar contratos zerados (Volume MWh = 0)", value=False)
            with col_flag3: flag_zerar_intercompany = st.toggle("🏢 Zerar InterCompany", value=False)

            base_original = base.copy()
            mask_intercompany = pd.Series(False, index=base.index)
            if flag_zerar_intercompany:
                base, mask_intercompany = aplicar_zerar_intercompany(base)

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

            _vol_mwm_num = pd.to_numeric(base["Volume MWm"], errors="coerce")
            _mask_valido_book = _vol_mwm_num.notna() & (base["Volume MWm"].astype(str).str.strip() != "-")
            _df_book = base[["Contrato CliqCCEE"]].copy()
            _df_book["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
            _soma_book = _df_book.groupby("Contrato CliqCCEE")["_vol_num"].transform("sum")
            base["Volume Book"] = _soma_book

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
                    
                    res_val = d_field.get(cod, "-")
                    
                    if pd.isna(res_val) or str(res_val).strip().lower() in ["nan", "none", ""]:
                        return "-"

                    if isinstance(res_val, str) and res_val != "-":
                        res_val_clean = res_val.replace(",", ".").strip()
                        try:
                            return float(res_val_clean)
                        except:
                            return res_val
                    return res_val

                base["Modulação Mínima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_min, idx_b_min, idx_a_min), axis=1)
                base["Modulação Máxima CCEE"] = base.apply(lambda r: buscar_campo_ccee(r, idx_m_max, idx_b_max, idx_a_max), axis=1)
                
                def buscar_tipo_ccee(row, dict_m, dict_b, dict_a):
                    cod = str(row["Contrato CliqCCEE"]).strip()
                    if cod in ["", "-", "None", "nan", "Verificar"]: return "-"
                    try: b_int = int(float(str(row["BOLETA"]).strip()))
                    except: b_int = -1
                    d_field = dict_a if b_int in BOLETAS_ACR else (dict_b if str(row["Parte"]).strip().upper() == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A." else dict_m)
                    res_tipo = d_field.get(cod, "-")
                    if pd.isna(res_tipo) or str(res_tipo).strip().lower() in ["nan", "none", ""]:
                        return "-"
                    return res_tipo
                    
                base["Modulação CCEE"]        = base.apply(lambda r: buscar_tipo_ccee(r, idx_m_tipo, idx_b_tipo, idx_a_tipo), axis=1)
            else:
                base["Modulação Mínima CCEE"] = "-"
                base["Modulação Máxima CCEE"] = "-"
                base["Modulação CCEE"]        = "-"

            _tol_mod = 1e-4
            
            base["Check Modulação Mínima"] = "-"
            _mod_min_cc = pd.to_numeric(base["Modulação Mínima CCEE"], errors="coerce")
            _mask_min_valid = _mask_calcular_min & _mod_min_cc.notna()
            
            _mask_ambos_traco_min = (base["Modulação Mínima"].astype(str).str.strip() == "-") & (base["Modulação Mínima CCEE"].astype(str).str.strip() == "-")
            base.loc[_mask_ambos_traco_min, "Check Modulação Mínima"] = "OK"

            if _mask_min_valid.any():
                _diff_min = pd.to_numeric(base.loc[_mask_min_valid, "Modulação Mínima"]) - _mod_min_cc.loc[_mask_min_valid]
                _mask_min_calc = _mask_min_valid & (base["Check Modulação Mínima"] != "OK")
                base.loc[_mask_min_calc, "Check Modulação Mínima"] = "OK"
                base.loc[_mask_min_calc & (_diff_min > _tol_mod), "Check Modulação Mínima"] = "Book maior"
                base.loc[_mask_min_calc & (_diff_min < -_tol_mod), "Check Modulação Mínima"] = "CCEE maior"

            base["Check Modulação Máxima"] = "-"
            _mod_max_cc = pd.to_numeric(base["Modulação Máxima CCEE"], errors="coerce")
            _mask_max_valid = _mask_calcular_max & _mod_max_cc.notna()
            
            _mask_ambos_traco_max = (base["Modulação Máxima"].astype(str).str.strip() == "-") & (base["Modulação Máxima CCEE"].astype(str).str.strip() == "-")
            base.loc[_mask_ambos_traco_max, "Check Modulação Máxima"] = "OK"

            if _mask_max_valid.any():
                _diff_max = pd.to_numeric(base.loc[_mask_max_valid, "Modulação Máxima"]) - _mod_max_cc.loc[_mask_max_valid]
                _mask_max_calc = _mask_max_valid & (base["Check Modulação Máxima"] != "OK")
                base.loc[_mask_max_calc, "Check Modulação Máxima"] = "OK"
                base.loc[_mask_max_calc & (_diff_max > _tol_mod), "Check Modulação Máxima"] = "Book maior"
                base.loc[_mask_max_calc & (_diff_max < -_tol_mod), "Check Modulação Máxima"] = "CCEE maior"

            base["Check Modulação"] = "-"
            _mask_ambos_traco_tipo = (base["Modulação WBC"].astype(str).str.strip() == "-") & (base["Modulação CCEE"].astype(str).str.strip() == "-")
            base.loc[_mask_ambos_traco_tipo, "Check Modulação"] = "OK"

            _mask_tipo_valid = _mask_tem_contrato & (~base["Modulação CCEE"].astype(str).str.strip().isin(["", "-", "None", "nan"]))
            if _mask_tipo_valid.any():
                _mask_div_tipo = base["Modulação WBC"].astype(str).str.strip().str.upper() != base["Modulação CCEE"].astype(str).str.strip().str.upper()
                _mask_tipo_calc = _mask_tipo_valid & (base["Check Modulação"] != "OK")
                base.loc[_mask_tipo_calc, "Check Modulação"] = "OK"
                base.loc[_mask_tipo_calc & _mask_div_tipo, "Check Modulação"] = "Divergente"

            _ordem_colunas = [
                "BOLETA", "Operação", "Varejista", "Tipo de Energia", "Parte", "Contraparte Razão Social", "Contraparte",
                "CP/LP", "CNPJ CONTRAPARTE", "Submercado", "Volume (MWh)", "Volume MWm", "CliqCCEE Paradigma",
                "Modulação WBC", "% Modulação Mínima", "Modulação Mínima", "Modulação Mínima CCEE", "Check Modulação Mínima",
                "% Modulação Máxima", "Modulação Máxima", "Modulação Máxima CCEE", "Check Modulação Máxima",
                "Modulação CCEE", "Check Modulação",
                "Contrato CliqCCEE mês anterior", "Vendedor", "Comprador", "Contrato CliqCCEE",
                "Volume Book", "Volume CCEE", "Check Volume", "Volume Global", "Volume Global CCEE", "Check Volume Global"
            ]
            base = base[[c for c in _ordem_colunas if c in base.columns]]

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

            _tol = 1e-6
            _vb = pd.to_numeric(base["Volume Book"], errors="coerce").fillna(0.0)
            _vc = pd.to_numeric(base["Volume CCEE"], errors="coerce").fillna(0.0)
            _diff_vol = _vb - _vc
            base["Check Volume"] = "OK"
            base.loc[_diff_vol > _tol, "Check Volume"] = "Book maior"
            base.loc[_diff_vol < -_tol, "Check Volume"] = "CCEE maior"

            _df_global = base[["Vendedor", "Comprador", "Submercado"]].copy()
            _df_global["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
            _soma_global = _df_global.groupby(["Vendedor", "Comprador", "Submercado"])["_vol_num"].transform("sum")
            base["Volume Global"] = _soma_global

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

            _vgb = pd.to_numeric(base["Volume Global"], errors="coerce").fillna(0.0)
            _vgc = pd.to_numeric(base["Volume Global CCEE"], errors="coerce").fillna(0.0)
            _diff_global = _vgb - _vgc
            base["Check Volume Global"] = "OK"
            base.loc[_diff_global > _tol, "Check Volume Global"] = "Book maior"
            base.loc[_diff_global < -_tol, "Check Volume Global"] = "CCEE maior"

            compras_net = base[base["Operação"] == "Compra"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"].sum().rename(columns={"Volume (MWh)": "Compra (MWh)"})
            vendas_net = base[base["Operação"] == "Venda"].groupby(["Parte", "Contraparte", "Submercado", "Tipo de Energia"], as_index=False)["Volume (MWh)"].sum().rename(columns={"Volume (MWh)": "Venda (MWh)"})
            nets = compras_net.merge(vendas_net, on=["Parte", "Contraparte", "Submercado", "Tipo de Energia"], how="inner")

            if flag_zerar_intercompany:
                n_ic = int(mask_intercompany.sum())
                st.info(f"🏢 **Zerar InterCompany ativo** — {n_ic} contrato(s) zerado(s) destacados em amarelo 🟡")

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

            col_f7, _, _ = st.columns(3)
            with col_f7: filtro_varejista = st.multiselect("Varejista", options=sorted(base_exibicao["Varejista"].unique()), default=[])

            if filtro_operacao: base_exibicao = base_exibicao[base_exibicao["Operação"].isin(filtro_operacao)]
            if filtro_status: base_exibicao = base_exibicao[base_exibicao["Contrato CliqCCEE"].astype(str).isin(filtro_status)]
            if filtro_submercado: base_exibicao = base_exibicao[base_exibicao["Submercado"].astype(str).isin(filtro_submercado)]
            if filtro_parte: base_exibicao = base_exibicao[base_exibicao["Parte"].astype(str).str.contains(filtro_parte, case=False, na=False)]
            if filtro_contraparte: base_exibicao = base_exibicao[base_exibicao["Contraparte"].astype(str).str.contains(filtro_contraparte, case=False, na=False)]
            if filtro_boleta: base_exibicao = base_exibicao[base_exibicao["BOLETA"].astype(str).str.contains(filtro_boleta, case=False, na=False)]
            if filtro_varejista: base_exibicao = base_exibicao[base_exibicao["Varejista"].isin(filtro_varejista)]

            if flag_ocultar_zerados: base_exibicao = base_exibicao[base_exibicao["Volume (MWh)"] != 0.0]

            base_exibicao["Volume (MWh)"] = base_exibicao["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            base_exibicao["Volume MWm"]   = base_exibicao["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            
            for c_format in ["Modulação Mínima", "Modulação Máxima", "Modulação Mínima CCEE", "Modulação Máxima CCEE"]:
                if c_format in base_exibicao.columns:
                    base_exibicao[c_format] = base_exibicao[c_format].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) and not pd.isna(x) else x)

            st.caption(f"{len(base_exibicao):,} registros encontrados")

            colunas_texto = [
                "BOLETA", "Operação", "Varejista", "Tipo de Energia", "Parte", "Contraparte Razão Social",
                "Contraparte", "CP/LP", "CNPJ CONTRAPARTE", "Submercado", "CliqCCEE Paradigma",
                "Contrato CliqCCEE mês anterior", "Contrato CliqCCEE", "Modulação WBC",
                "% Modulação Mínima", "% Modulação Máxima", "Modulação Mínima", "Modulação Máxima",
                "Modulação Mínima CCEE", "Modulação Máxima CCEE", "Check Modulação Mínima",
                "Check Modulação Máxima", "Modulação CCEE", "Check Modulação", "Vendedor", "Comprador"
            ]
            for col in colunas_texto:
                if col in base_exibicao.columns:
                    base_exibicao[col] = base_exibicao[col].astype(str).replace(["nan", "None", "NaN"], "-")

            _boletas_ef_set = st.session_state.get("boletas_efetivadas", set())
            _idx_intercompany = set(base.index[mask_intercompany].tolist()) if flag_zerar_intercompany else set()

            def _highlight_tabela(row):
                boleta_str = str(row.get("BOLETA", "")).strip()
                if boleta_str in _boletas_ef_set:
                    return ["background-color: #7B2D8B; color: white"] * len(row)
                if flag_zerar_intercompany and row.name in _idx_intercompany:
                    return ["background-color: #FFD700"] * len(row)
                if flag_mesmo_titular:
                    parte_r = str(row.get("Parte", "")).strip().upper()
                    contra_r = str(row.get("Contraparte Razão Social", "")).strip().upper()
                    if parte_r and contra_r and parte_r == contra_r:
                        return ["background-color: #FFD700"] * len(row)
                return [""] * len(row)

            styled = base_exibicao.style.apply(_highlight_tabela, axis=1)
            st.dataframe(styled, use_container_width=True, hide_index=True)

            base_download = base.copy()
            if flag_ocultar_zerados: base_download = base_download[base_download["Volume (MWh)"] != 0.0]
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer: base_download.to_excel(writer, sheet_name="Base Conferência", index=False)
            st.download_button("📥 Download Base Conferência", data=output.getvalue(), file_name="Base_Conferencia.xlsx")

            # ── RESUMO DE NETs (AQUI ESTAVA O ERRO DE DIGITAÇÃO DE COMPRAPOR) ───────────────────
            with st.expander("📋 Resumo de NETs", expanded=False):
                mes_ref_net = int(df["Mes"].dropna().iloc[0])
                text_horas_net = horas_mes.get(mes_ref_net, 744)

                _grp_key = ["Parte", "Contraparte", "Submercado", "Tipo de Energia"]

                _compras_mwm = (
                    base[base["Operação"] == "Compra"]
                    .groupby(_grp_key, as_index=False)["Volume MWm"]
                    .sum()
                    .rename(columns={"Volume MWm": "_compra_mwm"})
                )
                _vendas_mwm = (
                    base[base["Operação"] == "Venda"]
                    .groupby(_grp_key, as_index=False)["Volume MWm"]
                    .sum()
                    .rename(columns={"Volume MWm": "_venda_mwm"})
                )

                _nets_resumo = _compras_mwm.merge(_vendas_mwm, on=_grp_key, how="inner")
                _nets_resumo["_saldo"] = _nets_resumo["_compra_mwm"] - _nets_resumo["_venda_mwm"]

                def _quem_ajusta(saldo):
                    if saldo > 1e-9:
                        return "Contraparte"
                    elif saldo < -1e-9:
                        return "Parte"
                    return "Nenhum"

                _nets_resumo["_quem_ajusta_flag"] = _nets_resumo["_saldo"].apply(_quem_ajusta)

                def _resolver_ajustador(row):
                    flag = row["_quem_ajusta_flag"]
                    if flag == "Contraparte":
                        return row["Contraparte"]
                    elif flag == "Parte":
                        return row["Parte"]
                    return "Nenhum"

                _nets_resumo["_quem_ajusta_nome"] = _nets_resumo.apply(_resolver_ajustador, axis=1)
                _nets_resumo["_vol_ajustar"] = _nets_resumo["_saldo"].abs().where(
                    _nets_resumo["_quem_ajusta_flag"] != "Nenhum", 0.0
                )

                _dfs_ccee_net = []
                for _src in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
                    if _src is not None and not _src.empty:
                        _cols_need_net = [
                            "SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR",
                            "SUBMERCADO_ENTREGA", "MWmedio"
                        ]
                        if all(c in _src.columns for c in _cols_need_net):
                            _tmp_net = _src[_cols_need_net].copy()
                            _tmp_net["MWmedio"] = (
                                _tmp_net["MWmedio"].astype(str).str.strip()
                                .str.replace(",", ".", regex=False)
                            )
                            _tmp_net["MWmedio"] = pd.to_numeric(_tmp_net["MWmedio"], errors="coerce").fillna(0.0)
                            _dfs_ccee_net.append(_tmp_net)

                if _dfs_ccee_net:
                    _df_ccee_all = pd.concat(_dfs_ccee_net, ignore_index=True)
                    # CORRIGIDO: de 'SIGLA_PERFIL_COMPRAPOR' para 'SIGLA_PERFIL_COMPRADOR'
                    _idx_cliq = (
                        _df_ccee_all
                        .groupby(
                            ["SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR", "SUBMERCADO_ENTREGA"],
                            as_index=False
                        )["MWmedio"]
                        .sum()
                    )
                    _idx_cliq.rename(columns={"MWmedio": "_mwmedio_sum"}, inplace=True)
                else:
                    _idx_cliq = pd.DataFrame(
                        columns=["SIGLA_PERFIL_VENDEDOR", "SIGLA_PERFIL_COMPRADOR", "SUBMERCADO_ENTREGA", "_mwmedio_sum"]
                    )

                def _buscar_cliq_net(row):
                    parte = row["Parte"]
                    contraparte = row["Contraparte"]
                    sub = row["Submercado"]
                    tipo_en = row["Tipo de Energia"]

                    _mask_net = (
                        (base["Parte"] == parte) &
                        (base["Contraparte"] == contraparte) &
                        (base["Submercado"] == sub) &
                        (base["Tipo de Energia"] == tipo_en)
                    )
                    _boletas_net = base[_mask_net]

                    if _boletas_net.empty or _idx_cliq.empty:
                        return 0.0, 0.0

                    _compras_net = _boletas_net[_boletas_net["Operação"] == "Compra"]
                    _vendas_net = _boletas_net[_boletas_net["Operação"] == "Venda"]

                    compra_cliq = 0.0
                    venda_cliq = 0.0

                    for _, b_row in _compras_net.iterrows():
                        _v = str(b_row["Vendedor"]).strip()
                        _c = str(b_row["Comprador"]).strip()
                        if _v in ("-", "", "nan") or _c in ("-", "", "nan"):
                            continue
                        _match = _idx_cliq[
                            (_idx_cliq["SIGLA_PERFIL_VENDEDOR"] == _v) &
                            (_idx_cliq["SIGLA_PERFIL_COMPRADOR"] == _c) &
                            (_idx_cliq["SUBMERCADO_ENTREGA"] == sub)
                        ]
                        if not _match.empty:
                            compra_cliq += _match["_mwmedio_sum"].iloc[0]

                    for _, b_row in _vendas_net.iterrows():
                        _v = str(b_row["Vendedor"]).strip()
                        _c = str(b_row["Comprador"]).strip()
                        if _v in ("-", "", "nan") or _c in ("-", "", "nan"):
                            continue
                        _match = _idx_cliq[
                            (_idx_cliq["SIGLA_PERFIL_VENDEDOR"] == _v) &
                            (_idx_cliq["SIGLA_PERFIL_COMPRADOR"] == _c) &
                            (_idx_cliq["SUBMERCADO_ENTREGA"] == sub)
                        ]
                        if not _match.empty:
                            venda_cliq += _match["_mwmedio_sum"].iloc[0]

                    return compra_cliq, venda_cliq

                _cliq_results = _nets_resumo.apply(_buscar_cliq_net, axis=1, result_type="expand")
                _nets_resumo["_compra_cliq"] = _cliq_results[0]
                _nets_resumo["_venda_cliq"]  = _cliq_results[1]

                if "net_efetivados" not in st.session_state:
                    st.session_state["net_efetivados"] = {}

                boletas_efetivadas = set()
                _tol_cliq = 1e-6

                def _calcular_status(row):
                    if not row.get("_efetivado", False):
                        return "NET Não Efetivado"
                    vol_aj = row["_vol_ajustar"]
                    comp_cliq = row["_compra_cliq"]
                    vend_cliq = row["_venda_cliq"]
                    quem = row["_quem_ajusta_flag"]

                    if quem == "Nenhum":
                        return "OK"

                    if quem == "Contraparte":
                        cliq_aj = comp_cliq
                    else:
                        cliq_aj = vend_cliq

                    diff = abs(cliq_aj - vol_aj)
                    if diff < _tol_cliq:
                        return "OK"
                    elif cliq_aj == 0.0:
                        return "Não Ajustado"
                    elif cliq_aj < vol_aj - _tol_cliq:
                        return "Ajuste Parcial"
                    elif cliq_aj > vol_aj + _tol_cliq:
                        return "Volume Maior"
                    else:
                        return "Divergente"

                _col_headers = [
                    "Efetivado", "Parte", "Contraparte", "Submercado", "Tipo de Energia",
                    "Compra (MWm)", "Venda (MWm)", "Saldo NET (MWm)",
                    "Volume a Ajustar (MWm)", "Quem Ajusta",
                    "Compra Cliq (MWm)", "Venda Cliq (MWm)", "Status"
                ]

                _hdr = st.columns([0.5, 2, 2, 1.2, 1.5, 1.2, 1.2, 1.2, 1.5, 1.5, 1.5, 1.5, 1.5])
                for _ci, _ch in zip(_hdr, _col_headers):
                    _ci.markdown(f"**{_ch}**")

                st.markdown("---")

                for _i, _row in _nets_resumo.iterrows():
                    _net_key = (
                        str(_row["Parte"]) + "|" +
                        str(_row["Contraparte"]) + "|" +
                        str(_row["Submercado"]) + "|" +
                        str(_row["Tipo de Energia"])
                    )
                    _efetivado = st.session_state["net_efetivados"].get(_net_key, False)
                    _row["_efetivado"] = _efetivado
                    _status = _calcular_status(_row)

                    _status_cores = {
                        "OK": "🟢",
                        "Não Ajustado": "🔴",
                        "Ajuste Parcial": "🟡",
                        "Volume Maior": "🟠",
                        "Volume Menor": "🟠",
                        "Divergente": "🔴",
                        "NET Não Efetivado": "⚪",
                    }
                    _status_icon = _status_cores.get(_status, "⚪")

                    _saldo = _row["_saldo"]
                    _saldo_fmt = f"{_saldo:+.6f}"

                    _cols = st.columns([0.5, 2, 2, 1.2, 1.5, 1.2, 1.2, 1.2, 1.5, 1.5, 1.5, 1.5, 1.5])
                    _novo_ef = _cols[0].checkbox("", value=_efetivado, key=f"net_ef_{_net_key}")
                    if _novo_ef != _efetivado:
                        st.session_state["net_efetivados"][_net_key] = _novo_ef
                        st.rerun()

                    _cols[1].write(_row["Parte"])
                    _cols[2].write(_row["Contraparte"])
                    _cols[3].write(_row["Submercado"])
                    _cols[4].write(_row["Tipo de Energia"])
                    _cols[5].write(f"{_row['_compra_mwm']:.6f}")
                    _cols[6].write(f"{_row['_venda_mwm']:.6f}")
                    _cols[7].write(_saldo_fmt)
                    _cols[8].write(f"{_row['_vol_ajustar']:.6f}")
                    _cols[9].write(_row["_quem_ajusta_nome"])
                    _cols[10].write(f"{_row['_compra_cliq']:.6f}")
                    _cols[11].write(f"{_row['_venda_cliq']:.6f}")
                    _cols[12].write(f"{_status_icon} {_status}")

                    if _novo_ef:
                        _mask_ef = (
                            (base["Parte"] == _row["Parte"]) &
                            (base["Contraparte"] == _row["Contraparte"]) &
                            (base["Submercado"] == _row["Submercado"]) &
                            (base["Tipo de Energia"] == _row["Tipo de Energia"])
                        )
                        boletas_efetivadas.update(base[_mask_ef]["BOLETA"].astype(str).tolist())

                st.session_state["boletas_efetivadas"] = boletas_efetivadas

                st.markdown("---")
                output_nets = BytesIO()
                with pd.ExcelWriter(output_nets, engine="openpyxl") as writer:
                    _nets_resumo.to_excel(writer, sheet_name="Resumo de Nets", index=False)
                st.download_button(
                    label="📥 Download Resumo de NETs (.xlsx)",
                    data=output_nets.getvalue(),
                    file_name="Resumo_de_Nets.xlsx",
                    mime="application/vnd.ms-excel.sheet.macroEnabled.12"
                )

            # ── CONFERÊNCIA AVANÇADA DE DIVERGÊNCIAS ──
            if csvs_disponiveis:
                st.markdown("---")
                lista_divergencias = []
                lista_sem_match_nenhum = []

                for _, row in base.iterrows():
                    try: b_int = int(float(str(row["BOLETA"]).strip()))
                    except: b_int = -1

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

                    divs = []

                    if cod_encontrado:
                        v_c = d_v.get(cod_encontrado, '')
                        c_c = d_c.get(cod_encontrado, '')
                        s_c = d_s.get(cod_encontrado, '')

                        if v_b != v_c: divs.append(f"Divergência de Vendedor (Book={v_b} | CCEE={v_c})")
                        if c_b != c_c: divs.append(f"Divergência de Comprador (Book={c_b} | CCEE={c_c})")
                        if s_b != s_c: divs.append(f"Divergência de Submercado (Book={s_b} | CCEE={s_c})")

                    check_mod_min = str(row.get("Check Modulação Mínima", "-")).strip()
                    if check_mod_min in ("Book maior", "CCEE maior"):
                        divs.append(f"Modulação Mínima ({check_mod_min})")

                    check_mod_max = str(row.get("Check Modulação Máxima", "-")).strip()
                    if check_mod_max in ("Book maior", "CCEE maior"):
                        divs.append(f"Modulação Máxima ({check_mod_max})")

                    check_mod_tipo = str(row.get("Check Modulação", "-")).strip()
                    if check_mod_tipo == "Divergente":
                        divs.append("Tipo de Modulação Divergente")

                    check_vol_contrato = str(row.get("Check Volume", "-")).strip()
                    if check_vol_contrato in ("Book maior", "CCEE maior"):
                        divs.append(f"Volume do Contrato ({check_vol_contrato})")

                    if not cod_encontrado and chave_esp not in d_ch.values():
                        status = "SEM_MATCH"
                        justificativa = "Contrato inexistente no CSV CCEE"
                    elif divs:
                        status = "ERRO"
                        justificativa = " | ".join(divs)
                    else:
                        status = "OK"
                        justificativa = None

                    if status in ("ERRO", "SEM_MATCH"):
                        item = {"Boleta": row["BOLETA"], "Vendedor": row["Vendedor"], "Comprador": row["Comprador"], "Mensagem": justificativa}
                        if status == "ERRO": lista_divergencias.append(item)
                        else: lista_sem_match_nenhum.append(item)

                df_divergencias = pd.DataFrame(lista_divergencias, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])
                df_sem_match_nenhum = pd.DataFrame(lista_sem_match_nenhum, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])

                st.subheader("❌ Contratos com Divergência (Cadastro ou Indicadores Incorretos)")
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
