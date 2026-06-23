import streamlit as st
import pandas as pd
import numpy as np
import zipfile
import io
import re

st.set_page_config(page_title="Validador de Contratos CCEE", layout="wide")

st.title("Validador de Contratos CCEE")
st.write("Versão atualizada com filtros nativos nas colunas e contadores agregados.")

# --- Upload dos Arquivos ---
uploaded_base = st.file_uploader("Suba o Relatório Base (RelPers_858)", type=["xlsx", "csv"])
uploaded_mes_anterior = st.file_uploader("Suba a planilha do Mês Anterior", type=["xlsx", "csv"])
uploaded_zips = st.file_uploader("Suba os arquivos ZIP da CCEE (Matrix e Bismut)", type=["zip"], accept_multiple_files=True)

if uploaded_base and uploaded_mes_anterior and uploaded_zips:
    
    # 1. Leitura do Relatório Base
    try:
        if uploaded_base.name.endswith('.csv'):
            df_base = pd.read_csv(uploaded_base, header=8, sep=None, engine='python')
        else:
            df_base = pd.read_excel(uploaded_base, header=8)
    except Exception as e:
        st.error(f"Erro ao ler o relatório base: {e}")
        st.stop()
        
    # 2. Leitura do Mês Anterior
    try:
        if uploaded_mes_anterior.name.endswith('.csv'):
            df_mes_ant = pd.read_csv(uploaded_mes_anterior, sep=None, engine='python')
        else:
            df_mes_ant = pd.read_excel(uploaded_mes_anterior)
        df_mes_ant['BOLETA'] = df_mes_ant['BOLETA'].astype(str).str.strip()
        df_mes_ant['Codigo_CCEE'] = df_mes_ant['Codigo_CCEE'].astype(str).str.strip()
        dict_mes_ant = dict(zip(df_mes_ant['BOLETA'], df_mes_ant['Codigo_CCEE']))
    except Exception as e:
        st.error(f"Erro ao ler a planilha do mês anterior: {e}")
        st.stop()

    # 3. Processamento dos ZIPs CCEE
    df_ccee_matrix = pd.DataFrame()
    df_ccee_bismut = pd.DataFrame()
    df_ccee_acr = pd.DataFrame()

    BOLETAS_ACR = ["134882", "134884", "134886", "134888", "134890", "134892", "134894", "134896", "134898", "134900"]

    for uploaded_zip in uploaded_zips:
        with zipfile.ZipFile(io.BytesIO(uploaded_zip.read())) as z:
            for file_name in z.namelist():
                if "parcela" in file_name.lower() or not file_name.endswith('.csv'):
                    continue
                
                try:
                    with z.open(file_name) as f:
                        content = f.read().decode('utf-8', errors='ignore')
                        # Remove a linha de separador comum da CCEE (sep=)
                        content_clean = re.sub(r'^sep=.*\n', '', content)
                        df_tmp = pd.read_csv(io.StringIO(content_clean), sep='\t')
                        
                        if df_tmp.empty:
                            continue
                            
                        # Padronização de colunas essenciais
                        if 'CODIGO_CONTRATO' in df_tmp.columns:
                            df_tmp['CODIGO_CONTRATO'] = df_tmp['CODIGO_CONTRATO'].astype(str).str.strip()
                        
                        # Roteamento dos arquivos baseado nos nomes padrões da CCEE
                        if "ccear_q" in file_name.lower():
                            df_ccee_acr = pd.concat([df_ccee_acr, df_tmp], ignore_index=True)
                        elif "124274" in file_name:  # Código identificador Newave Bismut
                            df_ccee_bismut = pd.concat([df_ccee_bismut, df_tmp], ignore_index=True)
                        elif "124261" in file_name:  # Código identificador Comercializadora Matrix
                            df_ccee_matrix = pd.concat([df_ccee_matrix, df_tmp], ignore_index=True)
                except Exception as e:
                    st.warning(f"Aviso ao processar {file_name}: {e}")

    # --- Construção da Estrutura Final ---
    df_final = pd.DataFrame()
    
    # Mapeamento e extração de colunas da Base
    df_final['BOLETA'] = df_base['Codigo_WBC'].astype(str).str.strip()
    df_final['Operação'] = df_base['Movimentacao'].astype(str).str.strip()
    df_final['Tipo de Energia'] = df_base['Fonte_Contrato'].astype(str).str.strip()
    df_final['Parte'] = df_base['Parte_razao_social'].astype(str).str.strip()
    df_final['Contraparte Razão Social'] = df_base['Contraparte_razao_social'].astype(str).str.strip()
    df_final['CliqCCEE Paradigma'] = df_base['Codigo_CCEE'].astype(str).str.strip()
    
    # Aplicação da regra de correção de energia na Coluna 3 (Tipo de Energia) caso a Parte seja UFV JACARANDA 1
    df_final.loc[df_final['Parte'] == "UFV JACARANDA 1", 'Tipo de Energia'] = "Incentivada-I5"

    # Preenche coluna de histórico do mês anterior
    df_final['Contrato CliqCCEE mês anterior'] = df_final['BOLETA'].map(dict_mes_ant).fillna('-')

    # --- Lógica de Busca e Validação CCEE ---
    status_list = []
    contrato_ccee_list = []

    # Preparação das chaves de cruzamento CCEE (Retirando espaços extras)
    for df_ccee in [df_ccee_matrix, df_ccee_bismut, df_ccee_acr]:
        if not df_ccee.empty:
            for col in ['SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA']:
                if col in df_ccee.columns:
                    df_ccee[col] = df_ccee[col].astype(str).str.strip()

    for idx, row in df_final.iterrows():
        boleta = row['BOLETA']
        parte = row['Parte']
        cliq_paradigma = row['CliqCCEE Paradigma']
        status_resolvido = False
        contrato_encontrado = "-"

        # Regra 1: Validação de Boletas Fixas no CCEAR (ACR)
        if boleta in BOLETAS_ACR and not df_ccee_acr.empty:
            # Filtra correspondência aproximada ou direta
            match = df_ccee_acr[df_ccee_acr['CODIGO_CONTRATO'] == cliq_paradigma]
            if not match.empty:
                contrato_encontrado = match.iloc[0]['CODIGO_CONTRATO']
                status_list.append("OK")
                contrato_ccee_list.append(contrato_encontrado)
                continue

        # Seleção do lote da CCEE com base na comercializadora/parte
        if parte == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.":
            df_trabalho = df_ccee_bismut
        else:
            df_trabalho = df_ccee_matrix

        if df_trabalho.empty:
            status_list.append("Verificar")
            contrato_ccee_list.append("-")
            continue

        # Regra 2: Cruzamento por código direto (Paradigma ou Mês Anterior)
        for alvo in [cliq_paradigma, row['Contrato CliqCCEE mês anterior']]:
            if alvo and alvo != '-':
                match = df_trabalho[df_trabalho['CODIGO_CONTRATO'] == alvo]
                if not match.empty:
                    contrato_encontrado = alvo
                    status_resolvido = True
                    break

        if status_resolvido:
            status_list.append("OK")
            contrato_ccee_list.append(contrato_encontrado)
        else:
            status_list.append("Verificar")
            contrato_ccee_list.append("-")

    df_final['Contrato CliqCCEE Corrente'] = contrato_ccee_list
    df_final['Status'] = status_list

    # --- Indicadores Visuais (Cards/Balões no topo) ---
    total_contratos = len(df_final)
    total_compras = len(df_final[df_final['Operação'].str.upper() == 'COMPRA'])
    total_vendas = len(df_final[df_final['Operação'].str.upper() == 'VENDA'])

    col1, col2, col3 = st.columns(3)
    col1.metric(label="Total de Contratos", value=total_contratos)
    col2.metric(label="Contratos de Compra 📥", value=total_compras)
    col3.metric(label="Contratos de Venda 📤", value=total_vendas)

    st.markdown("---")

    # --- Visualização com Filtro Nativo (Opção 2) ---
    st.subheader("Visualização e Filtro dos Dados")
    st.info("💡 Clique nos cabeçalhos das colunas para ordenar ou use os ícones de lupa/filtro para refinar as linhas.")
    
    # O st.dataframe exibe a tabela e permite filtros dinâmicos na tela por padrão
    st.dataframe(
        df_final, 
        use_container_width=True,
        hide_index=True
    )

    # --- Exportação ---
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Base Conferência V18')
    towrite.seek(0)
    
    st.download_button(
        label="📥 Baixar Resultados para Excel",
        data=towrite,
        file_name="Base_Conferencia_Processada.xlsx",
        mime="application/vnd.ms-excel"
    )
else:
    st.info("Aguardando o upload de todos os arquivos obrigatórios para iniciar o cruzamento.")
