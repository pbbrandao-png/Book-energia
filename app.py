# APP_BOOK_ENERGIA_V21 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V19: + Ajuste Manual via Planilha | Separação de Erros de Divergência e Sem Match Nenhum nos CSVs
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro
# V21: + Liberação das colunas para ajuste interativo + Remoção total de rateios (Auto-referência)

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
    if "Editado Manualmente" in row.index and row["Editado Manualmente"] is True:
        return ["background-color: #D6EAF8"] * len(row)
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
arquivo_ajuste_manual = st.file_uploader("Selecione a planilha de Ajuste Manual", type=["xlsx"])
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

        mapa_ajuste_manual_paradigma = {}
        mapa_ajuste_manual_contraparte = {}
        mapa_ajuste_manual_vendedor = {}
        mapa_ajuste_manual_comprador = {}
        mapa_ajuste_manual_contrato = {}
        
        if arquivo_ajuste_manual is not None:
            try:
                df_aj_manual = pd.read_excel(arquivo_ajuste_manual)
                if "BOLETA" in df_aj_manual.columns:
                    df_aj_manual = df_aj_manual.dropna(subset=["BOLETA"])
                    df_aj_manual["BOLETA"] = df_aj_manual["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
                    if "CliqCCEE Paradigma" in df_aj_manual.columns:
                        mapa_ajuste_manual_paradigma = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["CliqCCEE Paradigma"]))
                    if "Contraparte" in df_aj_manual.columns:
                        mapa_ajuste_manual_contraparte = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["Contraparte"]))
                    if "Vendedor" in df_aj_manual.columns:
                        mapa_ajuste_manual_vendedor = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["Vendedor"]))
                    if "Comprador" in df_aj_manual.columns:
                        mapa_ajuste_manual_comprador = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["Comprador"]))
                    if "Contrato CliqCCEE" in df_aj_manual.columns:
                        mapa_ajuste_manual_contrato = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["Contrato CliqCCEE"]))
            except Exception as e_aj:
                st.error(f"Erro ao ler planilha de Ajuste Manual: {e_aj}")

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
        base["CliqCCEE Paradigma"]             = df["Codigo_CCEE"].astype(str).fillna("-")
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["% Modulação Mínima"]             = df["FlexLimite_modulacaoMin"].fillna("-")
        base["% Modulação Máxima"]             = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-").astype(str)
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str)
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str)
        base["Contrato CliqCCEE"]              = "-"
        base["Editado Manualmente"]            = False

        # Aplicação rápida dos ajustes de planilha externa
        boletas_str = base["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
        if mapa_ajuste_manual_paradigma:
            aj_p = boletas_str.map(mapa_ajuste_manual_paradigma)
            base["CliqCCEE Paradigma"] = aj_p.fillna(base["CliqCCEE Paradigma"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_p.notna()
        if mapa_ajuste_manual_contraparte:
            aj_c = boletas_str.map(mapa_ajuste_manual_contraparte)
            base["Contraparte"] = aj_c.fillna(base["Contraparte"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_c.notna()
        if mapa_ajuste_manual_vendedor:
            aj_v = boletas_str.map(mapa_ajuste_manual_vendedor)
            base["Vendedor"] = aj_v.fillna(base["Vendedor"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_v.notna()
        if mapa_ajuste_manual_comprador:
            aj_co = boletas_str.map(mapa_ajuste_manual_comprador)
            base["Comprador"] = aj_co.fillna(base["Comprador"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_co.notna()
        if mapa_ajuste_manual_contrato:
            aj_ct = boletas_str.map(mapa_ajuste_manual_contrato)
            base["Contrato CliqCCEE"] = aj_ct.fillna("-")
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_ct.notna()

        csvs_disponiveis = any([not df_ccee_matrix.empty, not df_ccee_bismut.empty, not df_ccee_acr.empty])

        if csvs_disponiveis:
            BISMUT_NOME_UPPER = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."
            
            def calcular_contrato_cliqccee_fast(row):
                if str(row.get("Contrato CliqCCEE", "-")).strip() not in ["", "-", "None", "nan"]:
                    if row.get("Editado Manualmente") and str(row["Contrato CliqCCEE"]) != "Verificar":
                        return str(row["Contrato CliqCCEE"])
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

            # Só roda para quem não foi forçado fixo pelo Excel anterior
            mask_auto = base["Contrato CliqCCEE"] == "-"
            if mask_auto.any():
                base.loc[mask_auto, "Contrato CliqCCEE"] = base[mask_auto].apply(calcular_contrato_cliqccee_fast, axis=1).astype(str)
        else:
            base["Contrato CliqCCEE"] = "-"

        if "base_editada" not in st.session_state:
            st.session_state["base_editada"] = base.copy()
        else:
            df_atual = base.copy()
            df_salvo = st.session_state["base_editada"]
            df_salvo = df_salvo[df_salvo["Editado Manualmente"] == True]
            if not df_salvo.empty:
                df_atual.set_index("BOLETA", inplace=True)
                df_salvo.set_index("BOLETA", inplace=True)
                
                df_salvo = df_salvo[~df_salvo.index.duplicated(keep='first')]
                colunas_comuns = df_atual.columns.intersection(df_salvo.columns)
                for col in colunas_comuns:
                    df_atual[col] = df_atual[col].astype(object)
                    df_salvo[col] = df_salvo[col].astype(object)
                
                df_atual.update(df_salvo)
                df_atual.reset_index(inplace=True)
                base = df_atual
            else:
                base = df_atual
            st.session_state["base_editada"] = base.copy()

        base = st.session_state["base_editada"]

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
            "Contrato CliqCCEE mês anterior", "Vendedor", "Comprador", "Contrato CliqCCEE", "Editado Manualmente",
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

            # ── CONFIGURAÇÃO DE COLUNAS: LIBERADO VENDEDOR, COMPRADOR E CONTRATO CLIQCCEE ──
            col_config = {
                "BOLETA": st.column_config.Column(disabled=True), "Operação": st.column_config.Column(disabled=True),
                "Tipo de Energia": st.column_config.Column(disabled=True), "Parte": st.column_config.Column(disabled=True),
                "Contraparte Razão Social": st.column_config.Column(disabled=True), "Contraparte": st.column_config.TextColumn(disabled=False),
                "CP/LP": st.column_config.Column(disabled=True), "CNPJ CONTRAPARTE": st.column_config.Column(disabled=True),
                "Submercado": st.column_config.Column(disabled=True), "Volume (MWh)": st.column_config.Column(disabled=True),
                "Volume MWm": st.column_config.Column(disabled=True), "CliqCCEE Paradigma": st.column_config.TextColumn(disabled=False),
                "Modulação WBC": st.column_config.Column(disabled=True), 
                "% Modulação Mínima": st.column_config.Column(disabled=True), "Modulação Mínima": st.column_config.Column(disabled=True), "Modulação Mínima CCEE": st.column_config.Column(disabled=True), "Check Modulação Mínima": st.column_config.Column(disabled=True),
                "% Modulação Máxima": st.column_config.Column(disabled=True), "Modulação Máxima": st.column_config.Column(disabled=True), "Modulação Máxima CCEE": st.column_config.Column(disabled=True), "Check Modulação Máxima": st.column_config.Column(disabled=True),
                "Modulação CCEE": st.column_config.Column(disabled=True), "Check Modulação": st.column_config.Column(disabled=True),
                "Contrato CliqCCEE mês anterior": st.column_config.TextColumn(disabled=True),
                "Vendedor": st.column_config.TextColumn(disabled=False),     # Liberado para ajuste interativo
                "Comprador": st.column_config.TextColumn(disabled=False),    # Liberado para ajuste interativo
                "Contrato CliqCCEE": st.column_config.TextColumn(disabled=False), # Liberado para ajuste interativo
                "Editado Manualmente": st.column_config.Column(disabled=True),
                "Volume Book": st.column_config.Column(disabled=True), "Volume CCEE": st.column_config.Column(disabled=True),
                "Check Volume": st.column_config.Column(disabled=True), "Volume Global": st.column_config.Column(disabled=True),
                "Volume Global CCEE": st.column_config.Column(disabled=True), "Check Volume Global": st.column_config.Column(disabled=True),
            }

            if flag_mesmo_titular:
                styled = base_exibicao.style.apply(highlight_mesmo_titular, axis=1)
                base_editada_df = st.data_editor(styled, use_container_width=True, hide_index=True, column_config=col_config, key="editor_base")
            else:
                base_editada_df = st.data_editor(base_exibicao, use_container_width=True, hide_index=True, column_config=col_config, key="editor_base")

            if st.session_state.get("editor_base") and st.session_state["editor_base"].get("edited_rows"):
                edicoes = st.session_state["editor_base"]["edited_rows"]
                indices_exibicao = base_exibicao.index.tolist()

                for idx_str, alteracoes in edicoes.items():
                    idx = int(idx_str)
                    idx_real = indices_exibicao[idx]
                    base.loc[idx_real, "Editado Manualmente"] = True
                    for col, val in alteracoes.items():
                        base.loc[idx_real, col] = str(val)
                    
                    # Se alterou Vendedor/Comprador e não definiu manualmente o contrato, força revalidação rápida
                    if csvs_disponiveis and "Contrato CliqCCEE" not in alteracoes:
                        base.loc[idx_real, "Contrato CliqCCEE"] = str(calcular_contrato_cliqccee_fast(base.loc[idx_real]))
                        
                st.session_state["base_editada"] = base.copy()
                st.rerun()

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
