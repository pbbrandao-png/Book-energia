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

st.sidebar.subheader("2. Base de Apoio")
arquivo_anterior = st.sidebar.file_uploader("Upload Mês Anterior (CliqCCEE)", type=['xlsx'], key="anterior")

st.title("📑 Book de Energia")

# 4. PROCESSAMENTO DA BASE ANTERIOR
dict_mes_anterior = {}
if arquivo_anterior:
    try:
        df_apoio = pd.read_excel(arquivo_anterior)
        dict_mes_anterior = pd.Series(df_apoio.iloc[:, 1].values, index=df_apoio.iloc[:, 0].values).to_dict()
        st.sidebar.success("✅ Base Mês Anterior carregada!")
    except Exception as e:
        st.sidebar.error(f"Erro na base anterior: {e}")

# 5. PROCESSAMENTO DA BASE PRINCIPAL
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

        df_conferencia = df_bruto[[col_boleta]].drop_duplicates().sort_values(by=col_boleta)
        df_conferencia.columns = ['Boleta']

        def buscar(coluna_alvo):
            return df_bruto.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)[coluna_alvo]

        # Preenchimento de dados
        df_conferencia['Operação'] = df_conferencia['Boleta'].map(buscar(col_operacao)).astype(str)
        
        trad_en = {"Incentivada-50%": "Incentivada-I5", "Incentivada-CQ50%": "Incentivada-CQ5", "Incentivada-100%": "Incentivada-I1", "Incentivada-0%": "Incentivada-I0", "Convencional": "Convencional"}
        df_conferencia['Tipo de Energia'] = df_conferencia['Boleta'].map(buscar(col_energia)).replace(trad_en)
        df_conferencia['Parte'] = df_conferencia['Boleta'].map(buscar(col_parte)).astype(str)
        df_conferencia['Contraparte'] = df_conferencia['Boleta'].map(buscar(col_contraparte))
        df_conferencia['CNPJ Contraparte'] = df_conferencia['Boleta'].map(buscar(col_cnpj)).apply(formatar_cnpj)
        
        # Volume calculado e arredondado
        v_mwh = df_conferencia['Boleta'].map(buscar(col_volume_mwh))
        h_mes = df_conferencia['Boleta'].map(buscar(col_horas_mes))
        df_conferencia['Volume MWm'] = (v_mwh / h_mes).fillna(0).round(4)
        
        df_conferencia['CliqCCEE Paradigma'] = df_conferencia['Boleta'].map(buscar(col_cliq_para))
        df_conferencia['Modulação WBC'] = df_conferencia['Boleta'].map(buscar(col_mod_wbc)).apply(limpar_modulacao).astype(str)
        df_conferencia['Modulação Mínima'] = df_conferencia['Boleta'].map(buscar(col_mod_min))
        df_conferencia['Modulação Máxima'] = df_conferencia['Boleta'].map(buscar(col_mod_max))
        df_conferencia['Contrato CliqCCEE mês anterior'] = df_conferencia['Boleta'].map(dict_mes_anterior).fillna("-")

        # --- 6. ÁREA DE FILTROS ---
        st.write("### Filtros da Tabela")
        f1, f2, f3, f4 = st.columns(4)

        with f1:
            op_list = ["Todos"] + sorted(df_conferencia['Operação'].unique().tolist())
            op_selected = st.selectbox("Operação", op_list)

        with f2:
            parte_list = ["Todos"] + sorted(df_conferencia['Parte'].unique().tolist())
            parte_selected = st.selectbox("Parte", parte_list)

        with f3:
            vol_list = ["Todos"] + sorted(df_conferencia['Volume MWm'].unique().astype(str).tolist(), key=float)
            vol_selected = st.selectbox("Volume MWm Específico", vol_list)

        with f4:
            mod_list = ["Todos"] + sorted(df_conferencia['Modulação WBC'].unique().tolist())
            mod_selected = st.selectbox("Modulação", mod_list)

        # CHECKBOX PARA REMOVER ZERADOS (Novo!)
        remover_zerados = st.checkbox("Ocultar Volumes Zerados (0.0000)", value=False)

        # --- APLICANDO OS FILTROS ---
        df_filtrado = df_conferencia.copy()
        
        if remover_zerados:
            df_filtrado = df_filtrado[df_filtrado['Volume MWm'] != 0]
        
        if op_selected != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Operação'] == op_selected]
        
        if parte_selected != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Parte'] == parte_selected]
            
        if vol_selected != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Volume MWm'] == float(vol_selected)]
            
        if mod_selected != "Todos":
            df_filtrado = df_filtrado[df_filtrado['Modulação WBC'] == mod_selected]

        # --- 7. EXIBIÇÃO ---
        st.markdown("---")
        colunas_exibicao = [
            'Boleta', 'Operação', 'Tipo de Energia', 'Parte', 
            'Contraparte', 'CNPJ Contraparte', 'Volume MWm', 
            'CliqCCEE Paradigma', 'Modulação WBC', 'Modulação Mínima', 
            'Modulação Máxima', 'Contrato CliqCCEE mês anterior'
        ]

        st.dataframe(
            df_filtrado[colunas_exibicao], 
            hide_index=True, 
            column_config={
                'Volume MWm': st.column_config.NumberColumn("Volume MWm", format="%.4f"),
                'Modulação Mínima': st.column_config.NumberColumn("Mod. Mínima", format="%.2f"),
                'Modulação Máxima': st.column_config.NumberColumn("Mod. Máxima", format="%.2f")
            },
            use_container_width=True
        )
        
        st.caption(f"Mostrando {len(df_filtrado)} de {len(df_conferencia)} operações.")
        
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
else:
    st.info("Aguardando upload dos arquivos.")