import streamlit as st
import pandas as pd
import re
from datetime import datetime

# 1. CONFIGURAÇÃO DA PÁGINA
st.set_page_config(layout="wide", page_title="Book de Energia")

# ─────────────────────────────────────────────────────────────────────────────
# LISTA DE DESTAQUE: CÓDIGOS CCEAR Q (AJUSTE AQUI)
# ─────────────────────────────────────────────────────────────────────────────
CODIGOS_CCEAR_Q_FORCADOS = [
    "2813298", "2813299", "2813300", "2813301", "2813302", "2813303", 
    "2813304", "2813305", "4159778", "4159779", "4159780", "4686267", 
    "4686268", "4686269", "4686270"
]

# 2. FUNÇÕES DE APOIO
def formatar_cnpj(cnpj):
    if pd.isna(cnpj) or cnpj == "": return ""
    apenas_numeros = re.sub(r'\D', '', str(cnpj)).zfill(14)
    return f"{apenas_numeros[:2]}.{apenas_numeros[2:5]}.{apenas_numeros[5:8]}/{apenas_numeros[8:12]}-{apenas_numeros[12:]}"

def tratar_chave(valor):
    if pd.isna(valor): return ""
    s = str(valor).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

# 3. INTERFACE LATERAL
st.sidebar.title("Configurações")
meses_nomes = ["Janeiro", "Fevereiro", "Marco", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
mes_nome_sel = st.sidebar.selectbox("Mês", meses_nomes, index=datetime.now().month - 1)
ano_sel_val = st.sidebar.selectbox("Ano", [str(a) for a in range(2024, 2031)], index=2)
mes_num_sel = meses_nomes.index(mes_nome_sel) + 1

st.sidebar.subheader("Regras de Volume")
# Mudamos o texto para ficar claro que agora a regra é ZERAR
aplicar_zerar_intra = st.sidebar.checkbox("Zerar Volume Intraportifólio", value=True)
aplicar_zerar_empresas = st.sidebar.checkbox("Zerar Volume Entre Empresas", value=True)
filtro_zeros_total = st.sidebar.checkbox("Ocultar Linhas com Volume 0", value=False)

# Uploads
st.sidebar.markdown("---")
arquivo_subido = st.sidebar.file_uploader("1. Contratos Aprovados (Excel)", type=['xlsx', 'xlsm'])

# 4. PROCESSAMENTO
if arquivo_subido:
    try:
        df_base = pd.read_excel(arquivo_subido, sheet_name='Contratos_Selecionados')
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        
        # Filtro por Mês
        df_conferencia = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_conferencia.empty:
            col_boleta = df_base.columns[0]
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            # Mapeamento
            df_conferencia['Operacao'] = df_conferencia.iloc[:, 1].astype(str)
            df_conferencia['Parte'] = df_conferencia.iloc[:, 62].astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia.iloc[:, 2].astype(str).str.strip()
            
            # Cálculo Inicial de Volume MWm
            v_mwh = pd.to_numeric(df_conferencia.iloc[:, 20], errors='coerce').fillna(0)
            h_mes = pd.to_numeric(df_conferencia.iloc[:, 15], errors='coerce').fillna(1)
            df_conferencia['Volume MWm'] = (v_mwh / h_mes).round(6)

            # ─────────────────────────────────────────────────────────────────
            # APLICAÇÃO DA REGRA: ZERAR VOLUME (NÃO REMOVER LINHA)
            # ─────────────────────────────────────────────────────────────────
            if aplicar_zerar_intra:
                mask_intra = df_conferencia['Operacao'].str.contains('INTRAPORTFOLIO', case=False, na=False)
                df_conferencia.loc[mask_intra, 'Volume MWm'] = 0
            
            if aplicar_zerar_empresas:
                mask_empresas = df_conferencia['Operacao'].str.contains('ENTRE EMPRESAS', case=False, na=False)
                df_conferencia.loc[mask_empresas, 'Volume MWm'] = 0
                
            # Se o usuário quiser esconder tudo que resultou em zero (opcional)
            if filtro_zeros_total:
                df_conferencia = df_conferencia[df_conferencia['Volume MWm'] != 0]

            # --- BALÕES DE QUANTIDADE (CONTADOR DE OPERAÇÕES) ---
            qtd_compra = len(df_conferencia[df_conferencia['Operacao'].str.contains('Compra', case=False, na=False)])
            qtd_venda = len(df_conferencia[df_conferencia['Operacao'].str.contains('Venda', case=False, na=False)])
            total_ops = len(df_conferencia)

            st.title(f"Book de Energia - {mes_nome_sel}/{ano_sel_val}")
            m1, m2, m3 = st.columns(3)
            m1.metric("Qtd. Operações Compra", f"{qtd_compra}")
            m2.metric("Qtd. Operações Venda", f"{qtd_venda}")
            m3.metric("Total de Operações", f"{total_ops}")
            st.markdown("---")

            # Coluna de Contrato Cliq (Exemplo da sua base)
            df_conferencia['Contrato CliqCCEE'] = df_conferencia.iloc[:, 60].apply(tratar_chave)
            
            # Aplicação do Destaque CCEAR Q
            def status_ajustado(row):
                if row['Contrato CliqCCEE'] in CODIGOS_CCEAR_Q_FORCADOS:
                    return "AJUSTE VALIDADO"
                return "-"

            df_conferencia['Status Montante'] = df_conferencia.apply(status_ajustado, axis=1)

            # Exibição Final
            colunas_finais = [col_boleta, 'Operacao', 'Parte', 'Volume MWm', 'Contrato CliqCCEE', 'Status Montante', 'Razao Social']
            st.dataframe(df_conferencia[colunas_finais], use_container_width=True, hide_index=True)
            
        else:
            st.warning("Sem dados para este período.")
    except Exception as e:
        st.error(f"Erro: {e}")
