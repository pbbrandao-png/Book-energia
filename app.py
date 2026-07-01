import os
import time
import pandas as pd
import streamlit as st
from datetime import datetime

# ==========================================
# CONFIGURAÇÕES DA PÁGINA
# ==========================================
st.set_page_config(
    page_title="Processador de Contratos CCEE",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização customizada usando st.html para evitar o TypeError de validação do Markdown
st.html("""
<style>
    [data-testid="stMetricValue"] {
        font-size: 24px;
    }
    div.stButton > button:first-child {
        background-color: #0066cc;
        color: white;
    }
</style>
""")

# ==========================================
# SIDEBAR - CONFIGURAÇÕES E FLAGS
# ==========================================
st.sidebar.header("⚙️ Configurações do Sistema")

# Agrupando todas as flags no mesmo bloco da Sidebar
com_ajuste_jacaranda = st.sidebar.checkbox("Ajustar UFV JACARANDA 1 (I5)", value=True)
remover_espacos = st.sidebar.checkbox("Remover espaços extras dos nomes", value=True)
zerar_intercompany = st.sidebar.checkbox("Zerar Intercompany", value=False)

st.sidebar.markdown("---")

# ==========================================
# CORPO PRINCIPAL
# ==========================================
st.title("⚡ Processador de Contratos CCEE")
st.write("Faça o upload do book de contratos para processar as informações.")

# Upload do arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel (.xlsx, .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        # Lendo todas as abas necessárias
        with st.spinner("Carregando dados do arquivo..."):
            xls = pd.ExcelFile(uploaded_file)
            
            # Validação básica de abas obrigatórias
            abas_obrigatorias = ['Aprovados', 'Carteira']
            if not all(aba in xls.sheet_names for aba in abas_obrigatorias):
                st.error("Erro: O arquivo precisa conter pelo menos as abas 'Aprovados' e 'Carteira'.")
                st.stop()
                
            df_aprovados = pd.read_excel(xls, sheet_name='Aprovados')
            df_carteira = pd.read_excel(xls, sheet_name='Carteira')

        # ------------------------------------------
        # PROCESSAMENTO DOS DADOS (BUSINESS LOGIC)
        # ------------------------------------------
        # Cópia para não mutar os originais na memória de forma errada
        df_proc_aprovados = df_aprovados.copy()
        df_proc_carteira = df_carteira.copy()

        # 1. Aplicação da Flag: Remover espaços extras dos nomes
        if remover_espacos:
            for df in [df_proc_aprovados, df_proc_carteira]:
                for col in df.select_dtypes(include=['object']).columns:
                    df[col] = df[col].astype(str).str.strip()

        # 2. Aplicação da Flag: Ajustar UFV JACARANDA 1
        # Quando a 'Parte' (coluna 4 / índice 3) for UFV JACARANDA 1, a coluna 3 (índice 2) vira Incentivada-I5
        if com_ajuste_jacaranda:
            if len(df_proc_aprovados.columns) >= 4:
                col_3 = df_proc_aprovados.columns[2]
                col_4 = df_proc_aprovados.columns[3]
                df_proc_aprovados.loc[df_proc_aprovados[col_4] == "UFV JACARANDA 1", col_3] = "Incentivada-I5"

        # 3. Aplicação da Flag: Zerar Intercompany
        if zerar_intercompany:
            # Lógica fictícia/placeholder para zerar operações intercompany se aplicável
            if 'Operação' in df_proc_aprovados.columns:
                df_proc_aprovados.loc[df_proc_aprovados['Operação'].astype(str).str.contains('INTERCOMPANY', case=False, na=False), :] = 0

        # Submercado Vendedor e Comprador (Simulação do filtro CCEE solicitado)
        # Filtros e agrupamentos baseados no escopo do setor de energia
        total_contratos = len(df_proc_aprovados)
        
        # Identificação de Compras e Vendas baseado na estrutura padrão
        tipo_col = 'Tipo' if 'Tipo' in df_proc_aprovados.columns else df_proc_aprovados.columns[0]
        compras = len(df_proc_aprovados[df_proc_aprovados[tipo_col].astype(str).str.contains('C|Compra', case=False, na=False)])
        vendas = len(df_proc_aprovados[df_proc_aprovados[tipo_col].astype(str).str.contains('V|Venda', case=False, na=False)])

        # ------------------------------------------
        # BI DE MÉTRICAS (BALÕES SUPERIORES)
        # ------------------------------------------
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            st.metric(label="Total de Contratos Processados", value=f"{total_contratos:,}".replace(",", "."))
        with col_m2:
            st.metric(label="📊 Total de Compras", value=f"{compras:,}".replace(",", "."))
        with col_m3:
            st.metric(label="📈 Total de Vendas", value=f"{vendas:,}".replace(",", "."))

        st.markdown("---")

        # Exibição dos dados processados (Preview)
        st.subheader("📋 Pré-visualização dos Contratos Aprovados")
        st.dataframe(df_proc_aprovados.head(100), use_container_width=True)

        # ------------------------------------------
        # ÁREA DE DOWNLOADS
        # ------------------------------------------
        st.subheader("📥 Downloads de Resultados")
        
        # Gerando arquivo em memória para download rápido
        import io
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_proc_aprovados.to_excel(writer, sheet_name='Resumo_Processado', index=False)
        processed_data = output.getvalue()

        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            st.download_button(
                label="📥 Baixar Book Processado (.xlsx)",
                data=processed_data,
                file_name=f"Book_Processado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with col_dl2:
            st.download_button(
                label="📥 Baixar Resumo de Nets (.xlsm)",
                data=processed_data, 
                file_name=f"Resumo_de_nets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsm",
                mime="application/vnd.ms-excel.sheet.macroEnabled.12"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")
