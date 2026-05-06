# --- DENTRO DA SEÇÃO 5. CARREGAMENTO DOS DADOS ---

if get_file_id(arquivo_pendencias) != st.session_state['fid_pendencias']:
    st.session_state['fid_pendencias'] = get_file_id(arquivo_pendencias)
    if arquivo_pendencias:
        try:
            # Carrega a base de pendências
            df_pend = pd.read_excel(arquivo_pendencias)
            
            # Identifica as colunas pelo índice (E = 4, I = 8)
            col_razao_index = 4
            col_valor_index = 8
            
            # Cria uma cópia para trabalhar
            df_temp = df_pend.iloc[:, [col_razao_index, col_valor_index]].copy()
            df_temp.columns = ['razao_orig', 'valor_pendente']
            
            # Limpa a razão social para o "match" e converte valor para numérico
            df_temp['razao_limpa'] = df_temp['razao_orig'].apply(limpar_str)
            df_temp['valor_pendente'] = pd.to_numeric(df_temp['valor_pendente'], errors='coerce').fillna(0)
            
            # --- LÓGICA DE SOMA (SOMASE) ---
            # Agrupa por razão social limpa e soma os valores
            df_agrupado = df_temp.groupby('razao_limpa')['valor_pendente'].sum().reset_index()
            
            # Transforma em dicionário para busca rápida
            st.session_state['dict_pendencias'] = pd.Series(
                df_agrupado.valor_pendente.values, 
                index=df_agrupado.razao_limpa.values
            ).to_dict()
            
            st.success("Base de Pendências Financeiras carregada e somada com sucesso!")
        except Exception as e:
            st.error(f"Erro ao processar base de pendências: {e}")
            st.session_state['dict_pendencias'] = {}

# --- DENTRO DA SEÇÃO 6. PROCESSAMENTO ---
# (Onde a coluna é criada no df_conferencia)

# Mapeamento de Pendência Financeira (Trazendo a soma já calculada)
df_conferencia['razao_limpa'] = df_conferencia['Razao Social'].apply(limpar_str)
df_conferencia['Pendência Financeira'] = df_conferencia['razao_limpa'].map(st.session_state['dict_pendencias']).fillna(0.0)
