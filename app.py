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
    "12850", "12851", "12852", "12853", "12854", "12855", "12856", "12857", "12858", "12859",
    "12860", "12861", "12862", "12863", "12864", "12865", "12866", "12867", "12868", "12869",
    "12870", "12871", "12872", "12873", "12874", "12875", "12876", "12877", "12878", "12879",
    "12880", "12881", "12882", "12883", "12884", "12885", "12886", "12887", "12888", "12889",
    "12890", "12891", "12892", "12893", "12894", "12895", "12896", "12897", "12898", "12899",
    "12900", "12901", "12902", "12903", "12904", "12905", "12906", "12907", "12908", "12909",
    "12910", "12911", "12912", "12913", "12914", "12915", "12916", "12917", "12918", "12919",
    "12920", "12921", "12922", "12923", "12924", "12925", "12926", "12927", "12928", "12929",
    "12930", "12931", "12932", "12933", "12934", "12935", "12936", "12937", "12938", "12939",
    "12940", "12941", "12942", "12943", "12944", "12945", "12946", "12947", "12948", "12949",
    "12950", "12951", "12952", "12953", "12954", "12955", "12956", "12957", "12958", "12959",
    "12960", "12961", "12962", "12963", "12964", "12965", "12966", "12967", "12968", "12969",
    "12970", "12971", "12972", "12973", "12974", "12975", "12976", "12977", "12978", "12979",
    "12980", "12981", "12982", "12983", "12984", "12985", "12986", "12987", "12988", "12989",
    "12990", "12991", "12992", "12993", "12994", "12995", "12996", "12997", "12998", "12999",
    "13000", "13001", "13002", "13003", "13004", "13005", "13006", "13007", "13008", "13009",
    "13010", "13011", "13012", "13013", "13014", "13015", "13016", "13017", "13018", "13019",
    "13020", "13021", "13022", "13023", "13024", "13025", "13026", "13027", "13028", "13029",
    "13030", "13031", "13032", "13033", "13034", "13035", "13036", "13037", "13038", "13039",
    "13040", "13041", "13042", "13043", "13044", "13045", "13046", "13047", "13048", "13049",
    "13050", "13051", "13052", "13053", "13054", "13055", "13056", "13057", "13058", "13059",
    "13060", "13061", "13062", "13063", "13064", "13065", "13066", "13067", "13068", "13069",
    "13070", "13071", "13072", "13073", "13074", "13075", "13076", "13077", "13078", "13079",
    "13080", "13081", "13082", "13083", "13084", "13085", "13086", "13087", "13088", "13089",
    "13090", "13091", "13092", "13093", "13094", "13095", "13096", "13097", "13098", "13099",
    "13100", "13101", "13102", "13103", "13104", "13105", "13106", "13107", "13108", "13109",
    "13110", "13111", "13112", "13113", "13114", "13115", "13116", "13117", "13118", "13119",
    "13120", "13121", "13122", "13123", "13124", "13125", "13126", "13127", "13128", "13129",
    "13130", "13131", "13132", "13133", "13134", "13135", "13136", "13137", "13138", "13139",
    "13140", "13141", "13142", "13143", "13144", "13145", "13146", "13147", "13148", "13149",
    "13150", "13151", "13152", "13153", "13154", "13155", "13156", "13157", "13158", "13159",
    "13160", "13161", "13162", "13163", "13164", "13165", "13166", "13167", "13168", "13169",
    "13170", "13171", "13172", "13173", "13174", "13175", "13176", "13177", "13178", "13179",
    "13180", "13181", "13182", "13183", "13184", "13185", "13186", "13187", "13188", "13189",
    "13190", "13191", "13192", "13193", "13194", "13195", "13196", "13197", "13198", "13199",
    "13200", "13201", "13202", "13203", "13204", "13205", "13206", "13207", "13208", "13209",
    "13210", "13211", "13212", "13213", "13214", "13215", "13216", "13217", "13218", "13219",
    "13220", "13221", "13222", "13223", "13224", "13225", "13226", "13227", "13228", "13229",
    "13230", "13231", "13232", "13233", "13234", "13235", "13236", "13237", "13238", "13239",
    "13240", "13241", "13242", "13243", "13244", "13245", "13246", "13247", "13248", "13249",
    "13250", "13251", "13252", "13253", "13254", "13255", "13256", "13257", "13258", "13259",
    "13260"
}

# Dicionário de horas por mês comercial
horas_mes = {
    1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720,
    7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744
}

# Mapeamento de Razão Social para Código de Perfil CCEE da Contraparte
razao_para_perfil = {
    "MATRIX COMERCIALIZADORA DE ENERGIA ELETRICA S.A.": "MATRIX",
    "BISMUT COMERCIALIZADORA DE ENERGIA LTDA": "BISMUT",
    "MATRIX VAREJISTA COMERCIALIZADORA DE ENERGIA S.A.": "MATRIX VAR",
    "BISMUT VAREJISTA COMERCIALIZADORA DE ENERGIA LTDA.": "BISMUT VAR"
}

def extrair_csvs_de_zip(file_bytes):
    """Extrai arquivos CSV específicos de um arquivo ZIP em memória da CCEE de forma otimizada."""
    csv_data = {}
    with zipfile.ZipFile(BytesIO(file_bytes)) as z:
        for name in z.namelist():
            if "cceal_firme" in name and name.endswith(".csv"):
                csv_data["cceal_firme"] = pd.read_csv(z.open(name), sep=";", encoding="iso-8859-1")
            elif "ccear_q" in name and name.endswith(".csv"):
                csv_data["ccear_q"] = pd.read_csv(z.open(name), sep=";", encoding="iso-8859-1")
            elif "cbr_mercado_proprio" in name and name.endswith(".csv"):
                csv_data["cbr_mercado_proprio"] = pd.read_csv(z.open(name), sep=";", encoding="iso-8859-1")
    return csv_data

def processar_csv_cceal_firme(df):
    """Processa o dataframe do arquivo cceal_firme padronizando as colunas."""
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    colunas_uteis = ["Código do Contrato", "Sigla do Submercado Comprador", "Sigla do Submercado Vendedor", "Nome do Perfil do Vendedor", "Nome do Perfil do Comprador"]
    colunas_existentes = [col for col in colunas_uteis if col in df.columns]
    df = df[colunas_existentes].copy()
    df = df.rename(columns={
        "Código do Contrato": "Contrato CliqCCEE",
        "Sigla do Submercado Comprador": "Submercado Comprador CCEE",
        "Sigla do Submercado Vendedor": "Submercado Vendedor CCEE",
        "Nome do Perfil do Vendedor": "Perfil Vendedor CCEE",
        "Nome do Perfil do Comprador": "Perfil Comprador CCEE"
    })
    df["Contrato CliqCCEE"] = df["Contrato CliqCCEE"].astype(str).str.strip()
    return df

def processar_csv_ccear_q(df):
    """Processa o dataframe do arquivo ccear_q padronizando as colunas."""
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    colunas_uteis = ["Código do Contrato", "Submercado", "Nome do Perfil do Comprador"]
    colunas_existentes = [col for col in colunas_uteis if col in df.columns]
    df = df[colunas_existentes].copy()
    df = df.rename(columns={
        "Código do Contrato": "Contrato CliqCCEE",
        "Submercado": "Submercado Comprador CCEE",
        "Nome do Perfil do Comprador": "Perfil Comprador CCEE"
    })
    df["Contrato CliqCCEE"] = df["Contrato CliqCCEE"].astype(str).str.strip()
    df["Submercado Vendedor CCEE"] = "SE"
    df["Perfil Vendedor CCEE"] = "Geradora ACR"
    return df

def processar_csv_cbr_mercado_proprio(df):
    """Processa o dataframe do arquivo cbr_mercado_proprio padronizando as colunas."""
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = df.columns.str.strip()
    colunas_uteis = ["Código do Contrato", "Sigla do Submercado Comprador", "Sigla do Submercado Vendedor", "Nome do Perfil do Vendedor", "Nome do Perfil do Comprador"]
    colunas_existentes = [col for col in colunas_uteis if col in df.columns]
    df = df[colunas_existentes].copy()
    df = df.rename(columns={
        "Código do Contrato": "Contrato CliqCCEE",
        "Sigla do Submercado Comprador": "Submercado Comprador CCEE",
        "Sigla do Submercado Vendedor": "Submercado Vendedor CCEE",
        "Nome do Perfil do Vendedor": "Perfil Vendedor CCEE",
        "Nome do Perfil do Comprador": "Perfil Comprador CCEE"
    })
    df["Contrato CliqCCEE"] = df["Contrato CliqCCEE"].astype(str).str.strip()
    return df

st.set_page_config(page_title="Validador de Book de Energia", layout="wide")
st.title("Validador de Book de Energia vs CCEE")

st.sidebar.header("1. Upload de Arquivos")
book_file = st.sidebar.file_uploader("Upload do Book de Energia (Excel)", type=["xlsx", "xlsm"])
zip_matrix_file = st.sidebar.file_uploader("Upload do ZIP Matrix (CCEE)", type=["zip"])
zip_bismut_file = st.sidebar.file_uploader("Upload do ZIP Bismut (CCEE)", type=["zip"])

if book_file and (zip_matrix_file or zip_bismut_file):
    with st.spinner("Processando dados de entrada..."):
        df_book = pd.read_excel(book_file, sheet_name="Registros")
        df_book.columns = df_book.columns.str.strip()
        
        # Limpeza e padronização inicial do Book
        df_book["Boleta"] = df_book["Boleta"].astype(str).str.strip()
        df_book["Contrato CliqCCEE"] = df_book["Contrato CliqCCEE"].astype(str).str.strip()
        df_book["Parte"] = df_book["Parte"].astype(str).str.strip()
        df_book["Contraparte"] = df_book["Contraparte"].astype(str).str.strip()
        df_book["Contraparte Razão Social"] = df_book["Contraparte Razão Social"].astype(str).str.strip()
        df_book["Operação"] = df_book["Operação"].astype(str).str.upper().str.strip()
        
        # Aplicação da regra customizada de Fonte de Energia
        mask_jacaranda = df_book["Parte"] == "UFV JACARANDA 1"
        if "Fonte de energia" in df_book.columns:
            df_book.loc[mask_jacaranda, "Fonte de energia"] = "Incentivada-I5"
        elif df_book.shape[1] >= 3:
            df_book.iloc[mask_jacaranda, 2] = "Incentivada-I5"
        
        # Determinação do Perfil Esperado da Contraparte baseado na Razão Social
        df_book["Perfil Contraparte Esperado CCEE"] = df_book["Contraparte Razão Social"].map(razao_para_perfil).fillna("OUTROS")
        
        # Processamento dos arquivos CCEE da Matrix
        db_matrix = pd.DataFrame()
        if zip_matrix_file:
            csvs_matrix = extrair_csvs_de_zip(zip_matrix_file.read())
            df_cceal_m = processar_csv_cceal_firme(csvs_matrix.get("cceal_firme"))
            df_ccear_m = processar_csv_ccear_q(csvs_matrix.get("ccear_q"))
            df_cbr_m = processar_csv_cbr_mercado_proprio(csvs_matrix.get("cbr_mercado_proprio"))
            db_matrix = pd.concat([df_cceal_m, df_ccear_m, df_cbr_m], ignore_index=True).drop_duplicates(subset=["Contrato CliqCCEE"])
            
        # Processamento dos arquivos CCEE da Bismut
        db_bismut = pd.DataFrame()
        if zip_bismut_file:
            csvs_bismut = extrair_csvs_de_zip(zip_bismut_file.read())
            df_cceal_b = processar_csv_cceal_firme(csvs_bismut.get("cceal_firme"))
            db_bismut = df_cceal_b.drop_duplicates(subset=["Contrato CliqCCEE"])

    # Criação dos filtros e flags organizados na barra lateral
    st.sidebar.header("2. Filtros e Opções")
    
    # Opção para ocultar contratos intraportfólio juntada com as demais opções
    ocultar_intraportfólio = st.sidebar.checkbox("Ocultar Contratos Intraportfólio", value=False)
    ocultar_zerados = st.sidebar.checkbox("Ocultar Boletas com Volume ZERADO", value=False)
    
    lista_operacoes = ["TODOS"] + sorted(df_book["Operação"].dropna().unique().tolist())
    operacao_selecionada = st.sidebar.selectbox("Filtrar por Operação", lista_operacoes)
    
    lista_partes = ["TODOS"] + sorted(df_book["Parte"].dropna().unique().tolist())
    parte_selecionada = st.sidebar.selectbox("Filtrar por Parte", lista_partes)
    
    lista_cliq = ["TODOS"] + sorted(df_book["Contrato CliqCCEE"].dropna().unique().tolist())
    cliq_selecionado = st.sidebar.selectbox("Filtrar por Contrato CliqCCEE", lista_cliq)

    # Identificação do Perfil Interno com base na Parte (Matrix ou Bismut)
    def identificar_perfil_interno(parte_str):
        p_upper = parte_str.upper()
        if "BISMUT VAREJISTA" in p_upper or "BISMUT VAR" in p_upper:
            return "BISMUT VAR"
        elif "MATRIX VAREJISTA" in p_upper or "MATRIX VAR" in p_upper:
            return "MATRIX VAR"
        elif "BISMUT" in p_upper:
            return "BISMUT"
        else:
            return "MATRIX"

    df_book["Perfil Interno Esperado CCEE"] = df_book["Parte"].apply(identificar_perfil_interno)

    # Função de conciliação de contratos individuais de forma vetorizada
    def preencher_dados_ccee_vectorized(df):
        df["Submercado Vendedor CCEE"] = ""
        df["Submercado Comprador CCEE"] = ""
        df["Perfil Vendedor CCEE"] = ""
        df["Perfil Comprador CCEE"] = ""
        
        # Casos ACR (lista fixa de boletas)
        mask_acr = df["Boleta"].isin(BOLETAS_ACR)
        if mask_acr.any() and not db_matrix.empty:
            merged_acr = df[mask_acr][["Contrato CliqCCEE"]].merge(db_matrix, on="Contrato CliqCCEE", how="left")
            df.loc[mask_acr, ["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]] = merged_acr[["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]].values

        # Casos Não-ACR
        mask_nao_acr = ~mask_acr
        if mask_nao_acr.any():
            df_nao_acr = df[mask_nao_acr].copy()
            
            # Segmentação por Perfil Interno (Bismut vs Matrix)
            mask_bis = df_nao_acr["Perfil Interno Esperado CCEE"].str.contains("BISMUT")
            mask_mat = ~mask_bis
            
            # Cruzamento Bismut
            if mask_bis.any() and not db_bismut.empty:
                merged_bis = df_nao_acr[mask_bis][["Contrato CliqCCEE"]].merge(db_bismut, on="Contrato CliqCCEE", how="left")
                df_nao_acr.loc[mask_bis, ["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]] = merged_bis[["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]].values
            
            # Cruzamento Matrix
            if mask_mat.any() and not db_matrix.empty:
                merged_mat = df_nao_acr[mask_mat][["Contrato CliqCCEE"]].merge(db_matrix, on="Contrato CliqCCEE", how="left")
                df_nao_acr.loc[mask_mat, ["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]] = merged_mat[["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]].values
                
            df.loc[mask_nao_acr, ["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]] = df_nao_acr[["Submercado Vendedor CCEE", "Submercado Comprador CCEE", "Perfil Vendedor CCEE", "Perfil Comprador CCEE"]].values
            
        return df

    with st.spinner("Conciliando informações com a base CCEE..."):
        df_processado = preencher_dados_ccee_vectorized(df_book.copy())

    # Aplicação de validações de consistência lógica baseadas nas regras de mercado
    def validar_vendedor_comprador_vectorized(df):
        df["Validação Submercado Vendedor"] = "OK"
        df["Validação Submercado Comprador"] = "OK"
        df["Validação Perfil Vendedor"] = "OK"
        df["Validação Perfil Comprador"] = "OK"
        
        # Filtro de contratos não localizados na CCEE
        mask_not_found = df["Perfil Vendedor CCEE"].isna() | (df["Perfil Vendedor CCEE"] == "")
        df.loc[mask_not_found, ["Validação Submercado Vendedor", "Validação Submercado Comprador", "Validação Perfil Vendedor", "Validação Perfil Comprador"]] = "CONTRATO NÃO ENCONTRADO NA CCEE"
        
        mask_valid = ~mask_not_found
        if mask_valid.any():
            df_v = df[mask_valid].copy()
            
            # Validação para operações de COMPRA
            mask_compra = df_v["Operação"] == "COMPRA"
            if mask_compra.any():
                df_vc = df_v[mask_compra]
                
                # Validações de submercado e perfil para compras
                sub_v_err = df_vc["Submercado Vendedor"] != df_vc["Submercado Vendedor CCEE"]
                sub_c_err = df_vc["Submercado Comprador"] != df_vc["Submercado Comprador CCEE"]
                
                perf_v_err = df_vc["Perfil Vendedor"] != df_vc["Perfil Vendedor CCEE"]
                perf_c_err = df_vc["Perfil Interno Esperado CCEE"] != df_vc["Perfil Comprador CCEE"]
                
                df_v.loc[mask_compra, "Validação Submercado Vendedor"] = np.where(sub_v_err, "ERRO: Divergência Submercado Vendedor", "OK")
                df_v.loc[mask_compra, "Validação Submercado Comprador"] = np.where(sub_c_err, "ERRO: Divergência Submercado Comprador", "OK")
                df_v.loc[mask_compra, "Validação Perfil Vendedor"] = np.where(perf_v_err, "ERRO: Perfil Vendedor diverge da CCEE", "OK")
                df_v.loc[mask_compra, "Validação Perfil Comprador"] = np.where(perf_c_err, "ERRO: Perfil Comprador diverge da CCEE", "OK")
            
            # Validação para operações de VENDA
            mask_venda = df_v["Operação"] == "VENDA"
            if mask_venda.any():
                df_vv = df_v[mask_venda]
                
                # Validações de submercado e perfil para vendas
                sub_v_err = df_vv["Submercado Vendedor"] != df_vv["Submercado Vendedor CCEE"]
                sub_c_err = df_vv["Submercado Comprador"] != df_vv["Submercado Comprador CCEE"]
                
                perf_v_err = df_vv["Perfil Interno Esperado CCEE"] != df_vv["Perfil Vendedor CCEE"]
                perf_c_err = (df_vv["Perfil Contraparte Esperado CCEE"] != "OUTROS") & (df_vv["Perfil Contraparte Esperado CCEE"] != df_vv["Perfil Comprador CCEE"])
                
                df_v.loc[mask_venda, "Validação Submercado Vendedor"] = np.where(sub_v_err, "ERRO: Divergência Submercado Vendedor", "OK")
                df_v.loc[mask_venda, "Validação Submercado Comprador"] = np.where(sub_c_err, "ERRO: Divergência Submercado Comprador", "OK")
                df_v.loc[mask_venda, "Validação Perfil Vendedor"] = np.where(perf_v_err, "ERRO: Perfil Vendedor diverge da CCEE", "OK")
                df_v.loc[mask_venda, "Validação Perfil Comprador"] = np.where(perf_c_err, "ERRO: Perfil Comprador diverge da CCEE", "OK")
                
            df.loc[mask_valid, ["Validação Submercado Vendedor", "Validação Submercado Comprador", "Validação Perfil Vendedor", "Validação Perfil Comprador"]] = df_v[["Validação Submercado Vendedor", "Validação Submercado Comprador", "Validação Perfil Vendedor", "Validação Perfil Comprador"]].values
            
        return df

    df_resultado = validar_vendedor_comprador_vectorized(df_processado)

    # Filtros de visualização da tabela principal aplicados dinamicamente
    if operacao_selecionada != "TODOS":
        df_resultado = df_resultado[df_resultado["Operação"] == operacao_selecionada]
    if parte_selecionada != "TODOS":
        df_resultado = df_resultado[df_resultado["Parte"] == parte_selecionada]
    if cliq_selecionado != "TODOS":
        df_resultado = df_resultado[df_resultado["Contrato CliqCCEE"] == cliq_selecionado]
    if ocultar_zerados:
        df_resultado = df_resultado[df_resultado["Volume (MWh)"] != 0]
    if ocultar_intraportfólio:
        df_resultado = df_resultado[df_resultado["Parte"] != df_resultado["Contraparte"]]

    # Geração dos painéis de resumo executivo com os balões de contagem
    st.markdown("### Resumo Executivo das Operações")
    c1, c2, c3 = st.columns(3)
    
    total_contratos_filtro = len(df_resultado)
    compras_filtro = len(df_resultado[df_resultado["Operação"] == "COMPRA"])
    vendas_filtro = len(df_resultado[df_resultado["Operação"] == "VENDA"])
    
    c1.metric("Total de Contratos", f"{total_contratos_filtro}")
    c2.metric("Operações de Compra 🟢", f"{compras_filtro}")
    c3.metric("Operações de Venda 🔵", f"{vendas_filtro}")

    # Renderização da base consolidada com estilização condicional de alertas
    st.markdown("### Base Consolidada de Contratos")
    
    def aplicar_estilo_linha(row):
        estilos = [""] * len(row)
        # Highlight amarelo para operações de mesma origem e destino (Auto-referência / Portfólio)
        if str(row["Parte"]).strip() == str(row["Contraparte"]).strip():
            return ["background-color: #fdfd96; color: black;"] * len(row)
            
        # Highlight vermelho claro para erros críticos de identificação do contrato
        if row["Validação Perfil Vendedor"] == "CONTRATO NÃO ENCONTRADO NA CCEE":
            return ["background-color: #ffb3b3; color: black;"] * len(row)
            
        # Destaques específicos por célula em caso de divergências pontuais
        if row["Validação Submercado Vendedor"] != "OK":
            idx = row.index.get_loc("Submercado Vendedor")
            estilos[idx] = "background-color: #ffccff; color: black;"
        if row["Validação Submercado Comprador"] != "OK":
            idx = row.index.get_loc("Submercado Comprador")
            estilos[idx] = "background-color: #ffccff; color: black;"
        if row["Validação Perfil Vendedor"] != "OK":
            idx = row.index.get_loc("Perfil Vendedor")
            estilos[idx] = "background-color: #ffccff; color: black;"
        if row["Validação Perfil Comprador"] != "OK":
            idx = row.index.get_loc("Perfil Comprador")
            estilos[idx] = "background-color: #ffccff; color: black;"
        return estilos

    df_styled = df_resultado.style.apply(aplicar_estilo_linha, axis=1)
    st.dataframe(df_styled, use_container_width=True)

    # Geração dos relatórios de erros estruturados para auditoria
    st.markdown("---")
    st.markdown("### Relatórios de Divergências e Inconsistências")

    # Regra V20: Ignorar Intraportfólio e Zerados nas tabelas de erro de forma mandatória
    df_erros_base = df_resultado[
        (df_resultado["Parte"] != df_resultado["Contraparte"]) & 
        (df_resultado["Volume (MWh)"] != 0)
    ]

    err_nao_encontrados = df_erros_base[df_erros_base["Validação Perfil Vendedor"] == "CONTRATO NÃO ENCONTRADO NA CCEE"]
    err_submercados = df_erros_base[
        (df_erros_base["Validação Submercado Vendedor"].str.contains("ERRO")) |
        (df_erros_base["Validação Submercado Comprador"].str.contains("ERRO"))
    ]
    err_perfis = df_erros_base[
        (df_erros_base["Validação Perfil Vendedor"].str.contains("ERRO")) |
        (df_erros_base["Validação Perfil Comprador"].str.contains("ERRO"))
    ]

    tab1, tab2, tab3 = st.tabs([
        f"Contratos Não Encontrados ({len(err_nao_encontrados)})",
        f"Divergência de Submercado ({len(err_submercados)})",
        f"Divergência de Perfil/Agente ({len(err_perfis)})"
    ])

    with tab1:
        if not err_nao_encontrados.empty:
            st.dataframe(err_nao_encontrados[["Boleta", "Contrato CliqCCEE", "Parte", "Contraparte", "Operação", "Volume (MWh)"]], use_container_width=True)
        else:
            st.success("Nenhum contrato pendente de localização na CCEE.")

    with tab2:
        if not err_submercados.empty:
            st.dataframe(err_submercados[["Boleta", "Contrato CliqCCEE", "Parte", "Contraparte", "Submercado Vendedor", "Submercado Vendedor CCEE", "Submercado Comprador", "Submercado Comprador CCEE"]], use_container_width=True)
        else:
            st.success("Todos os submercados estão perfeitamente alinhados.")

    with tab3:
        if not err_perfis.empty:
            st.dataframe(err_perfis[["Boleta", "Contrato CliqCCEE", "Parte", "Contraparte", "Perfil Vendedor", "Perfil Vendedor CCEE", "Perfil Comprador", "Perfil Comprador CCEE", "Validação Perfil Vendedor", "Validação Perfil Comprador"]], use_container_width=True)
        else:
            st.success("Todos os perfis de agentes estão em conformidade.")

    # Módulo de Fechamento de Balanço de Volumes por Submercado e Contraparte (Agrupamento Dinâmico)
    st.markdown("---")
    st.markdown("### Fechamento de Balanço de Energia (Netting)")

    df_net = df_resultado.copy()
    if not df_net.empty:
        # Padronização de nomes para evitar problemas com espaços em branco (Ex: JC COMERCIO AL - APE GERAÇÃO)
        df_net["Parte"] = df_net["Parte"].str.strip().replace(r'\s+', ' ', regex=True)
        df_net["Contraparte"] = df_net["Contraparte"].str.strip().replace(r'\s+', ' ', regex=True)

        submercados_net = sorted(df_net["Submercado Vendedor"].dropna().unique().tolist())
        submercado_net_sel = st.selectbox("Selecione o Submercado para o Fechamento", submercados_net)

        df_sub = df_net[(df_net["Submercado Vendedor"] == submercado_net_sel) | (df_net["Submercado Comprador"] == submercado_net_sel)]

        partes_disponiveis = sorted(list(set(df_sub["Parte"].dropna().tolist() + df_sub["Contraparte"].dropna().tolist())))
        parte_net_sel = st.selectbox("Selecione a Empresa Base do Portfólio", partes_disponiveis)

        contrapartes_disponiveis = sorted([p for p in partes_disponiveis if p != parte_net_sel])
        contraparte_net_sel = st.selectbox("Selecione a Contraparte para Netting", contrapartes_disponiveis)

        # Filtros de escopo locais para a matriz de cruzamento direcionado de boletas
        mask_c1 = (df_sub["Parte"] == parte_net_sel) & (df_sub["Contraparte"] == contraparte_net_sel)
        mask_c2 = (df_sub["Parte"] == contraparte_net_sel) & (df_sub["Contraparte"] == parte_net_sel)
        df_pair = df_sub[mask_c1 | mask_c2].copy()

        if not df_pair.empty:
            def classificar_fluxo_netting(row):
                if row["Parte"] == parte_net_sel:
                    return "COMPRA" if row["Operação"] == "COMPRA" else "VENDA"
                else:
                    return "VENDA" if row["Operação"] == "COMPRA" else "COMPRA"

            df_pair["Fluxo_Netting"] = df_pair.apply(classificar_fluxo_netting, axis=1)

            compras = df_pair[df_pair["Fluxo_Netting"] == "COMPRA"]
            vendas = df_pair[df_pair["Fluxo_Netting"] == "VENDA"]

            compras_calc = compras.copy()
            vendas_calc = vendas.copy()
            compras_calc["Volume MWm"] = compras_calc.apply(lambda r: r["Volume (MWh)"] / horas_mes.get(int(r["Mes"]), 744) if pd.notna(r["Mes"]) and r["Volume (MWh)"] else 0, axis=1)
            vendas_calc["Volume MWm"] = vendas_calc.apply(lambda r: r["Volume (MWh)"] / horas_mes.get(int(r["Mes"]), 744) if pd.notna(r["Mes"]) and r["Volume (MWh)"] else 0, axis=1)

            compras["Volume (MWh)"] = compras["Volume (MWh)"].apply(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            compras["Volume MWm"] = compras_calc["Volume MWm"].apply(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            vendas["Volume (MWh)"] = vendas["Volume (MWh)"].apply(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            vendas["Volume MWm"] = vendas_calc["Volume MWm"].apply(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

            st.markdown("## COMPRAS")
            st.dataframe(compras[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)
            st.markdown("## VENDAS")
            st.dataframe(vendas[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)

            total_compra, total_venda = compras_calc["Volume (MWh)"].sum(), vendas_calc["Volume (MWh)"].sum()
            saldo = total_compra - total_venda
            total_compra_mwm, total_venda_mwm = compras_calc["Volume MWm"].sum(), vendas_calc["Volume MWm"].sum()
            mes_referencia = int(df_net["Mes"].dropna().iloc[0])
            saldo_mwm = saldo / horas_mes.get(mes_referencia, 744)

            ajuste = contraparte_net_sel if saldo > 0 else parte_net_sel if saldo < 0 else "ZERADO"
            resumo = pd.DataFrame({
                "Tipo": ["Compras", "Vendas", "Saldo"],
                "MWh": [f"{total_compra:.3f}", f"{total_venda:.3f}", f"{saldo:.3f}"],
                "MWm": [f"{total_compra_mwm:.6f}", f"{total_venda_mwm:.6f}", f"{saldo_mwm:.6f}"]
            })
            st.markdown("### Resumo do Netting")
            st.dataframe(resumo, hide_index=True)
            st.info(f"Para zerar o saldo deste submercado, a empresa **{ajuste}** deve realizar um ajuste de **{abs(saldo):.3f} MWh** (**{abs(saldo_mwm):.6f} MWm**).")
        else:
            st.warning("Nenhuma operação registrada entre as empresas selecionadas para este submercado.")
else:
    st.info("Aguardando o upload do Book de Energia (Excel) e de ao menos um dos arquivos ZIP da CCEE (Matrix ou Bismut) para iniciar a validação.")
