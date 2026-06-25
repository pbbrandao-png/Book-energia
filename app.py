# APP_BOOK_ENERGIA_V20 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V19: + Ajuste Manual via Planilha | Separação de Erros de Divergência e Sem Match Nenhum nos CSVs
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro

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

        horas_mes = {
            1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720,
            7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744
        }

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
            mapa_mes_anterior = dict(zip(df_mes_anterior["BOLETA"], df_mes_anterior["Codigo_CCEE"]))
        else:
            mapa_mes_anterior = {}

        mapa_ajuste_manual_vendedor = {}
        mapa_ajuste_manual_comprador = {}
        mapa_ajuste_manual_contrato = {}
        if arquivo_ajuste_manual is not None:
            try:
                df_aj_manual = pd.read_excel(arquivo_ajuste_manual)
                if "BOLETA" in df_aj_manual.columns:
                    df_aj_manual = df_aj_manual.dropna(subset=["BOLETA"])
                    df_aj_manual["BOLETA"] = df_aj_manual["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
                    if "VENDEDOR" in df_aj_manual.columns:
                        mapa_ajuste_manual_vendedor = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["VENDEDOR"]))
                    if "COMPRADOR" in df_aj_manual.columns:
                        mapa_ajuste_manual_comprador = dict(zip(df_aj_manual["BOLETA"], df_aj_manual["COMPRADOR"]))
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
        base["CliqCCEE Paradigma"]             = df["Codigo_CCEE"].fillna("-").astype(str)
        base["Modulação WBC"]                  = df["Tipo_de_modulacao"].astype(str).str.strip().map(mapa_modulacao).fillna(df["Tipo_de_modulacao"])
        base["% Modulação Mínima"]             = df["FlexLimite_modulacaoMin"].fillna("-")
        base["Modulação Mínima"]               = "-"
        base["Modulação Mínima CCEE"]          = "-"
        base["Check Modulação Mínima"]         = "-"
        base["% Modulação Máxima"]             = df["FlexLimite_modulacaoMax"].fillna("-")
        base["Modulação Máxima"]               = "-"
        base["Modulação Máxima CCEE"]          = "-"
        base["Check Modulação Máxima"]         = "-"
        base["Modulação CCEE"]                 = "-"
        base["Check Modulação"]                = "-"
        base["Contrato CliqCCEE mês anterior"] = base["BOLETA"].map(mapa_mes_anterior).fillna("-").astype(str)
        base["Vendedor"]                       = df["Sigla_CCEE_vendedor"].fillna("-").astype(str)
        base["Comprador"]                      = df["Sigla_CCEE_comprador"].fillna("-").astype(str)
        base["Editado Manualmente"]            = False

        # Aplicação rápida dos ajustes via map de chaves strings
        boletas_str = base["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
        if mapa_ajuste_manual_vendedor:
            aj_v = boletas_str.map(mapa_ajuste_manual_vendedor)
            base["Vendedor"] = aj_v.fillna(base["Vendedor"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_v.notna()
        if mapa_ajuste_manual_comprador:
            aj_c = boletas_str.map(mapa_ajuste_manual_comprador)
            base["Comprador"] = aj_c.fillna(base["Comprador"])
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_c.notna()
        if mapa_ajuste_manual_contrato:
            aj_ct = boletas_str.map(mapa_ajuste_manual_contrato)
            base["Contrato CliqCCEE"] = aj_ct.fillna("-")
            base["Editado Manualmente"] = base["Editado Manualmente"] | aj_ct.notna()
            # Registra no session state para não recalcular via regra CCEE o que foi forçado no Excel
            if "contratos_editados_diretamente" not in st.session_state:
                st.session_state["contratos_editados_diretamente"] = []
            for b_idx, b_val in enumerate(boletas_str):
                if b_val in mapa_ajuste_manual_contrato and b_val not in st.session_state["contratos_editados_diretamente"]:
                    st.session_state["contratos_editados_diretamente"].append(base.iloc[b_idx]["BOLETA"])

        mask_mesmo_titular = base["Parte"].str.strip().str.upper() == base["Contraparte Razão Social"].str.strip().str.upper()
        base.loc[mask_mesmo_titular, ["Volume (MWh)", "Volume MWm"]] = 0.0

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

            # Só calcula via automação CCEE as linhas que NÃO vieram preenchidas do ajuste manual
            if "Contrato CliqCCEE" not in base.columns:
                base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee_fast, axis=1).astype(str)
            else:
                mask_n_definido = base["Contrato CliqCCEE"] == "-"
                if mask_n_definido.any():
                    base.loc[mask_n_definido, "Contrato CliqCCEE"] = base[mask_n_definido].apply(calcular_contrato_cliqccee_fast, axis=1).astype(str)
        else:
            if "Contrato CliqCCEE" not in base.columns:
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
                # SOLUÇÃO DO ERRO: Remove índices duplicados do df_salvo para permitir o .update()
                df_salvo = df_salvo
