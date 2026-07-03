# APP_BOOK_ENERGIA_V22 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro
# V21: + Remoção total de rateios (Auto-referência)
# V22: + Identificação e Filtro de Varejistas (MATRIX VAR / BISMUT VAR) + Correção de Escopo de 'nets' + Correção de Sintaxe no rename

import streamlit as st
import pandas as pd
import zipfile
import numpy as np
import re
from io import BytesIO

# Configura o limite do Pandas Styler para evitar o erro de estouro de células devido ao aumento de colunas
pd.set_option("styler.render.max_elements", 2000000)

# Boletas que devem buscar no CSV ccear_q em vez do cceal_firme
BOLETAS_ACR = {
    122387, 122389, 122391, 122393, 122395, 122397, 122399, 122401,
    144795, 144797, 144799, 148084, 148088, 148090, 148092, 148518,
    149316, 149318, 149320, 149322, 149324,
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


def _extrair_codigo_ponto(valor):
    if pd.isna(valor):
        return None
    m = re.search(r'ACL\((.*?)\)', str(valor))
    return m.group(1).strip() if m else None


def _ler_planilha_modelagem_ativo(arquivo):
    nome = getattr(arquivo, "name", "") or ""
    engine = "xlrd" if nome.lower().endswith(".xls") else None

    try:
        bruto = pd.read_excel(arquivo, header=None, nrows=30, engine=engine)
    except Exception:
        arquivo.seek(0)
        bruto = pd.read_excel(arquivo, header=None, nrows=30)

    linha_cabecalho = None
    for i in range(len(bruto)):
        if bruto.iloc[i].astype(str).str.strip().eq('Nº do Ativo').any():
            linha_cabecalho = i
            break

    if linha_cabecalho is None:
        linha_cabecalho = 0

    arquivo.seek(0)
    try:
        return pd.read_excel(arquivo, header=linha_cabecalho, engine=engine)
    except Exception:
        arquivo.seek(0)
        return pd.read_excel(arquivo, header=linha_cabecalho)


def carregar_mapa_parcela_carga(arquivo_ponto, arquivo_boletas, arquivo_modelagem_ativo=None):
    mapa = {}
    try:
        if arquivo_ponto is None or arquivo_boletas is None:
            return mapa

        df_ponto = pd.read_excel(arquivo_ponto)
        df_boletas = pd.read_excel(arquivo_boletas)

        if 'Conteúdo Expressão Contábil Processada' not in df_ponto.columns or 'Cód. Parcela - Carga' not in df_ponto.columns:
            return mapa
        if 'Ponto de Medição' not in df_boletas.columns or 'Código' not in df_boletas.columns:
            return mapa

        dict_billing = dict(zip(
            df_boletas['Ponto de Medição'].astype(str).str.strip(),
            pd.to_numeric(df_boletas['Código'], errors='coerce')
        ))

        df_ponto = df_ponto.copy()
        df_ponto['_CODIGO_PONTO'] = df_ponto['Conteúdo Expressão Contábil Processada'].apply(_extrair_codigo_ponto)
        df_ponto['_BOLETA'] = df_ponto['_CODIGO_PONTO'].map(dict_billing)

        if arquivo_modelagem_ativo is not None and 'Nº Seq Ativo' in df_ponto.columns:
            try:
                df_modelagem = _ler_planilha_modelagem_ativo(arquivo_modelagem_ativo)
                if df_modelagem is not None and 'Nº do Ativo' in df_modelagem.columns and 'Status' in df_modelagem.columns:
                    ativos_concluidos = set(
                        pd.to_numeric(
                            df_modelagem.loc[df_modelagem['Status'].astype(str).str.strip() == 'Concluída', 'Nº do Ativo'],
                            errors='coerce'
                        ).dropna()
                    )
                    df_ponto = df_ponto[pd.to_numeric(df_ponto['Nº Seq Ativo'], errors='coerce').isin(ativos_concluidos)]
            except Exception as e:
                st.warning(f"Não foi possível aplicar o filtro da planilha Exportação Solicitação Modelagem Ativo: {e}")

        df_ponto = df_ponto.dropna(subset=['_BOLETA'])
        df_ponto = df_ponto.drop_duplicates(subset=['_BOLETA'], keep='first')

        mapa = dict(zip(
            pd.to_numeric(df_ponto['_BOLETA'], errors='coerce'),
            df_ponto['Cód. Parcela - Carga']
        ))
    except Exception:
        mapa = {}
    return mapa


def carregar_mapa_relpers_301(zip_relpers):
    mapa_situacao = {}
    mapa_pagamento = {}
    if zip_relpers is None:
        return mapa_situacao, mapa_pagamento
    try:
        with zipfile.ZipFile(zip_relpers) as zf:
            nome_xlsx = None
            for nome in zf.namelist():
                if nome.lower().endswith('.xlsx'):
                    nome_xlsx = nome
                    break
            if nome_xlsx is None:
                return mapa_situacao, mapa_pagamento
            with zf.open(nome_xlsx) as f:
                df_301 = pd.read_excel(BytesIO(f.read()))
        df_301.columns = df_301.columns.astype(str).str.strip()

        if 'Codigo_WBC' not in df_301.columns:
            return mapa_situacao, mapa_pagamento

        df_301 = df_301.dropna(subset=['Codigo_WBC'])
        df_301['Codigo_WBC'] = pd.to_numeric(df_301['Codigo_WBC'], errors='coerce')
        df_301 = df_301.dropna(subset=['Codigo_WBC'])
        df_301 = df_301.drop_duplicates(subset=['Codigo_WBC'], keep='last')

        if 'Situacao_ERP' in df_301.columns:
            mapa_situacao = dict(zip(df_301['Codigo_WBC'], df_301['Situacao_ERP']))
        if 'Data_Vencimento_ordem' in df_301.columns:
            mapa_pagamento = dict(zip(df_301['Codigo_WBC'], df_301['Data_Vencimento_ordem']))
    except Exception:
        mapa_situacao = {}
        mapa_pagamento = {}
    return mapa_situacao, mapa_pagamento


def carregar_mapa_situacao_pagamento(arquivo_faturamento):
    mapa_situacao = {}
    mapa_pagamento = {}
    try:
        if arquivo_faturamento is None:
            return mapa_situacao, mapa_pagamento

        try:
            df_fat = pd.read_excel(arquivo_faturamento, header=5, usecols="A:L")
        except Exception:
            arquivo_faturamento.seek(0)
            df_fat = pd.read_excel(arquivo_faturamento, header=5)
        df_fat.columns = df_fat.columns.astype(str).str.strip()
        df_fat = df_fat.loc[:, ~df_fat.columns.str.startswith("Unnamed")]

        if 'Boleta' not in df_fat.columns:
            return mapa_situacao, mapa_pagamento

        df_fat = df_fat.dropna(subset=['Boleta'])
        df_fat['Boleta'] = pd.to_numeric(df_fat['Boleta'], errors='coerce')
        df_fat = df_fat.dropna(subset=['Boleta'])

        if 'Saldo Parcela' in df_fat.columns:
            saldo_num = pd.to_numeric(df_fat['Saldo Parcela'], errors='coerce').fillna(0.0)
        else:
            saldo_num = pd.Series(0.0, index=df_fat.index)

        df_fat['_SITUACAO'] = np.where(saldo_num > 0, 'Em Aberto', 'Pago')
        df_fat = df_fat.drop_duplicates(subset=['Boleta'], keep='last')

        mapa_situacao = dict(zip(df_fat['Boleta'], df_fat['_SITUACAO']))

        if 'Data Vencimento' in df_fat.columns:
            mapa_pagamento = dict(zip(df_fat['Boleta'], df_fat['Data Vencimento']))
    except Exception:
        mapa_situacao = {}
        mapa_pagamento = {}
    return mapa_situacao, mapa_pagamento


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


def highlight_mesmo_titular(row):
    parte = str(row.get("Parte", "")).strip().upper()
    contraparte_rs = str(row.get("Contraparte Razão Social", "")).strip().upper()
    if parte and contraparte_rs and parte == contraparte_rs:
        return ["background-color: #FFD700"] * len(row)
    return [""] * len(row)


_PARTES_INTERCOMPANY = {
    "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.",
    "GET COMERCIALIZADORA DE ENERGIA S.A.",
    "ARGENTUM COMERCIALIZADORA DE ENERGIA LTDA.",
}


def aplicar_zerar_intercompany(base: pd.DataFrame):
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

pagina = st.sidebar.radio("Menu", ["Base Conferência", "Encontro Energético"])
st.sidebar.markdown("---")
flag_ocultar_intraportfolio = st.sidebar.toggle("🟡 Ocultar IntraPortifólio (Parte = Contraparte)", value=True)
st.sidebar.markdown("---")

st.title("📊 Book Energia")

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx"])
arquivo_mes_anterior = st.file_uploader("Selecione a planilha Mês Anterior", type=["xlsx"])
zip_matrix = st.file_uploader("Selecione o ZIP Matrix", type=["zip"])
zip_bismut = st.file_uploader("Selecione o ZIP Bismut", type=["zip"])
arquivo_ponto_medicao = st.file_uploader("Selecione a planilha Ponto de Medição - MATRIX", type=["xlsx", "xls"])
arquivo_boletas = st.file_uploader("Selecione a planilha Boletas", type=["xlsx", "xls"])
arquivo_modelagem_ativo = st.file_uploader("Selecione a planilha Exportação Solicitação Modelagem Ativo", type=["xlsx", "xls"])
arquivo_faturamento_aberto = st.file_uploader("Selecione a planilha Faturamento em Aberto", type=["xlsx", "xls"])
zip_relpers_301 = st.file_uploader("Selecione o ZIP RelPers 301 (Mapa Financeiro)", type=["zip"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=8)

        if flag_ocultar_intraportfolio and "Parte_razao_social" in df.columns and "Contraparte_razao_social" in df.columns:
            mask_rateio_interno = df["Parte_razao_social"].astype(str).str.strip().str.upper() == df["Contraparte_razao_social"].astype(str).str.strip().str.upper()
            df = df[~mask_rateio_interno].reset_index(drop=True)

        if "Codigo_WBC" in df.columns and "Numero_referencia_contrato" in df.columns and "Rateio" in df.columns:
            mask_rateio_duplicado = (df["Codigo_WBC"].astype(str).str.strip() == df["Numero_referencia_contrato"].astype(str).str.strip()) & (df["Rateio"].astype(str).str.strip().str.upper() == "SIM")
            df = df[~mask_rateio_duplicado].reset_index(drop=True)

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

        mapa_parcela_carga = carregar_mapa_parcela_carga(arquivo_ponto_medicao, arquivo_boletas, arquivo_modelagem_ativo)
        mapa_situacao_pagamento, mapa_pagamento = carregar_mapa_situacao_pagamento(arquivo_faturamento_aberto)

        mapa_situacao_301, mapa_pagamento_301 = carregar_mapa_relpers_301(zip_relpers_301)
        mapa_situacao_pagamento.update(mapa_situacao_301)
        mapa_pagamento.update(mapa_pagamento_301)

        csvs_matrix = extrair_csvs_zip(zip_matrix)
        csvs_bismut = extrair_csvs_zip(zip_bismut)

        df_ccee_matrix = combiner_dfs([csvs_matrix['cceal'], csvs_matrix['cbr']])
        df_ccee_bismut = combiner_dfs([csvs_bismut['cceal']])
        df_ccee_acr = combiner_dfs([csvs_matrix['ccear_q']])

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

        base["Parcela de Carga"] = pd.to_numeric(base["BOLETA"], errors="coerce").map(mapa_parcela_carga)
        base["Parcela de Carga"] = base["Parcela de Carga"].apply(lambda v: str(int(v)) if pd.notna(v) else "-")

        _boleta_num = pd.to_numeric(base["BOLETA"], errors="coerce")
        base["Situação pagamento"] = _boleta_num.map(mapa_situacao_pagamento).fillna("Pago")
        base["Pagamento"] = _boleta_num.map(mapa_pagamento)
        base["Pagamento"] = base["Pagamento"].apply(lambda v: v.strftime("%d/%m/%Y") if isinstance(v, (pd.Timestamp,)) else ("-" if pd.isna(v) else str(v)))

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

        _df_pgto = base[["Contrato CliqCCEE", "Situação pagamento"]].copy()
        _df_pgto["_vol_num"] = _vol_mwm_num.where(_mask_valido_book, 0.0)
        _df_pgto["_vol_pago"] = _df_pgto["_vol_num"].where(_df_pgto["Situação pagamento"] == "Pago", 0.0)
        _soma_pago = _df_pgto.groupby("Contrato CliqCCEE")["_vol_pago"].transform("sum")
        base["SITUAÇÃO PGTO"] = np.where(np.isclose(_soma_pago, base["Volume Book"], atol=1e-6), "Pago", "-")

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

        _valores_divergentes_mod = ("Book maior", "CCEE maior")
        base["Limites Modulação"] = "OK"
        _mask_limites_verificar = (
            base["Check Modulação Mínima"].isin(_valores_divergentes_mod) |
            base["Check Modulação Máxima"].isin(_valores_divergentes_mod)
        )
        base.loc[_mask_limites_verificar, "Limites Modulação"] = "Verificar"

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

    except Exception as e:
        st.error(f"Erro ao processar os dados da RelPers: {e}")

    # ──────────────────────────────────────────────────────────────────────────────
    # O restante do código de processamento e tabelas continua inalterado daqui para baixo...
