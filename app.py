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
