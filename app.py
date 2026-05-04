import streamlit as st
import pandas as pd
import re

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# 2. FUNÇÕES DE APOIO
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "":
        return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj))
    apenas_numeros = apenas_numeros.zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def limpar_modulacao(texto):
    if pd.isna(texto): return ""
    t = str(texto).upper()
    if "FLAT" in t: return "Flat"
    if "CARGA" in t: return "Carga"
    if "DECLARADO" in t or "INFORMADO" in t: return "Declarado"
    if "GERA" in t: return "Geração"
    return texto

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")

st.sidebar.subheader("1. Base do Mês Atual")
arquivo_subido = st.sidebar.file_uploader("Upload da Base Bruta (Excel)", type=['xlsx', 'xlsm'], key="atual")

st.sidebar.subheader("2. Base de Apoio (CliqCCEE)")
arquivo_anterior = st.sidebar.file_uploader("Upload Mês Anterior", type=['xlsx'], key="anterior")

st.sidebar.subheader("3. Relatório de Pessoas")
arquivo_pessoas = st.sidebar.file_uploader("Upload RelPers_858 (4).xlsx", type=['xlsx'], key="pessoas")

st.title("📑 Book de Energia")

# 4. PROCESSAMENTO DA BASE ANTERIOR
dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_apoio = pd.read_excel(arquivo_anterior)
        # Garantindo que a chave do dicionário seja string para o de/para
        df_apoio.iloc[:, 0] = df_apoio.iloc[:, 0].astype(str).str.strip()
        dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()
        st.sidebar.success("✅ Base Mês Anterior carregada!")
    except Exception as e:
        st.sidebar.error(f"Erro na base anterior: {e}")

# 5. PROCESSAMENTO DA BASE DE PESSOAS (RelPers_858)
dict_vendedor = {}
dict_comprador = {}
if arquivo_pessoas:
    try:
        df_pers = pd.read_excel(arquivo_pessoas)
        # IMPORTANTE: Convertendo Boleta (Col D / Index 3) para String para evitar o N/A
        df_pers.iloc[:, 3] = df_pers.iloc[:, 3].astype(str).str.strip()
        
        dict_comprador = pd.Series(df_pers.iloc[:, 1].values, index=df_pers.iloc[:, 3].values).to_dict()
        dict_vendedor = pd.Series(df_pers.iloc[:, 2].values, index=df_pers.iloc[:, 3].values).to_dict()
        st.sidebar.success("✅ Relatório de Pessoas carregado!")
    except Exception as e:
        st.sidebar.error(f"Erro no Relatório de Pessoas: {e}")

# 6. PROCESSAMENTO DA BASE PRINCIPAL
if arquivo_subido:
    try:
        df_bruto = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        
        # Mapeamento de Colunas
        col_boleta = df_bruto.columns[0]
        col_operacao = df_bruto.columns[1]
        col_cnpj = df_bruto.columns[4]
        col_energia = df_bruto.columns[5]
        col_contraparte = df_bruto.columns[6]
        col_horas_mes = df_bruto.columns[15]
        col_volume_mwh = df_bruto.columns[20]
        col_mod_min = df_bruto.columns[28]
        col_mod_max = df_bruto.columns[29]
        col_cliq_para = df_bruto.columns[60]
        col_parte = df_bruto.columns[62]
        col_mod_wbc = df_bruto.columns[63]

        # Criando a DF de conferência e limpando a Boleta
        df_conferencia = df_bruto[[col_boleta]].drop_duplicates()
        df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].astype(str).str.strip()
        df_conferencia = df_conferencia.sort_values(by='Boleta_Key')

        # Função de busca interna
        df_temp_busca = df_bruto.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

        # Preenchimento de dados
        df_conferencia['Operação'] = df_conferencia[col_boleta].map(df_temp_busca[col_operacao]).astype(str)
        
        # Buscando Comprador e Vendedor usando a chave String
        df_conferencia['Comprador'] = df_conferencia['Boleta_Key'].map(dict_comprador).fillna("Não Encontrado")
        df_conferencia['Vendedor'] = df_conferencia['Boleta_Key'].map(dict_vendedor).fillna("Não Encontrado")

        trad_en = {"Incentivada-50%": "Incentivada-I5", "Incentivada-CQ50%": "Incentivada-CQ5", "Incentivada-100%": "Incentivada-I1", "Incentivada-0%": "Incentivada-I0", "Convencional": "Convencional"}
        df_conferencia['Tipo de Energia'] = df_conferencia[col_boleta].map(df_temp_busca[col_energia]).replace(trad_en)
        df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_temp_busca[col_parte]).astype(str)
        df_conferencia['Contraparte'] = df_conferencia[col_boleta].map(df_temp_busca[col_contraparte])
        df_conferencia['CNPJ Contraparte'] = df_conferencia[col_boleta].map(df_temp_busca[col_cnpj]).apply(formatar_cnpj)
        
        v_mwh = df_conferencia[col_boleta].map(df_temp_busca[col_volume_mwh])
        h_mes = df_conferencia[col_boleta].map(df_temp_busca[col_horas_mes])
        df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(4)
        
        df_conferencia['CliqCCEE Paradigma'] = df_conferencia[col_boleta].map(df_temp_busca[col_cliq_para])
        df_conferencia['Modulação WBC'] = df_conferencia[col_boleta].map(df_temp_busca[col_mod_wbc]).apply(limpar_modulacao)
        df_conferencia['Modulação Mínima'] = df_conferencia[col_boleta].map(df_temp_busca[col_mod_min])
        df_conferencia['Modulação Máxima'] = df_conferencia[col_boleta].map(df_temp_busca[col_mod_max])
        df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta_Key'].map(dict_mes_anterior).fillna("-")

        # --- 7. ÁREA DE FILTROS ---
        st.write("### Filtros da Tabela")
        f1, f2, f3, f4 = st.columns(4)

        with f1:
            op_selected = st.selectbox("Operação", ["Todos"] + sorted(df_conferencia['Operação'].unique().tolist()))
        with f2:
            parte_selected = st.selectbox("Parte", ["Todos"] + sorted(df_conferencia['Parte'].unique().tolist()))
        with f3:
            vol_list = ["Todos"] + sorted([str(v) for v in df_conferencia['Volume MWm'].unique()], key=float)
            vol_selected = st.selectbox("Volume MWm Específico", vol_list)
        with f4:
            mod_selected = st.selectbox("Modulação", ["Todos"] + sorted(df_conferencia['Modulação WBC'].unique().tolist()))

        remover_zerados = st.checkbox("Ocultar Volumes Zerados (0.0000)", value=False)

        # Filtros
        df_filtrado = df_conferencia.copy()
        if remover_zerados: df_filtrado = df_filtrado[df_filtrado['Volume MWm'] != 0]
        if op_selected != "Todos": df_filtrado = df_filtrado[df_filtrado['Operação'] == op_selected]
        if parte_selected != "Todos": df_filtrado = df_filtrado[df_filtrado['Parte'] == parte_selected]
        if vol_selected != "Todos": df_filtrado = df_filtrado[df_filtrado['Volume MWm'] == float(vol_selected)]
        if mod_selected != "Todos": df_filtrado = df_filtrado[df_filtrado['Modulação WBC'] == mod_selected]

        # --- 8. EXIBIÇÃO E ORDEM DE COLUNAS ---
        # DEFINA A ORDEM AQUI (O Python seguirá exatamente esta lista)
        colunas_exibicao = [
            col_boleta,       # Boleta Original
            'Operação', 
            'Comprador',      # Agora no começo
            'Vendedor',       # Agora no começo
            'Tipo de Energia', 
            'Parte', 
            'Contraparte', 
            'CNPJ Contraparte', 
            'Volume MWm', 
            'CliqCCEE Paradigma', 
            'Modulação WBC', 
            'Modulação Mínima', 
            'Modulação Máxima', 
            'Contrato CliqCCEE mês anterior'
        ]

        st.markdown("---")
        st.dataframe(
            df_filtrado[colunas_exibicao], 
            hide_index=True, 
            column_config={
                'Volume MWm': st.column_config.NumberColumn("Volume MWm", format="%.4f"),
            },
            use_container_width=True
        )
        
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Aguardando upload dos arquivos.")
