# 6. PROCESSAMENTO DA TABELA
if st.session_state['df_bruto'] is not None:
    try:
        df_base = st.session_state['df_bruto'].copy()
        col_mes = df_base.columns[14]
        df_base[col_mes] = pd.to_numeric(df_base[col_mes], errors='coerce')
        df_filtrada = df_base[df_base[col_mes] == mes_num_sel].copy()

        if not df_filtrada.empty:
            col_boleta = df_base.columns[0]
            df_conferencia = df_filtrada[[col_boleta]].drop_duplicates()
            df_conferencia['Boleta_Key'] = df_conferencia[col_boleta].apply(tratar_chave)
            
            # Criamos o lookup garantindo que as colunas AC (28) e AD (29) estejam lá
            df_lookup = df_filtrada.drop_duplicates(subset=[col_boleta]).set_index(col_boleta)

            # --- MAPEAMENTO DAS COLUNAS ---
            df_conferencia['Operacao'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[1]]).astype(str)
            df_conferencia['Parte'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[62]]).astype(str).str.strip()
            df_conferencia['Razao Social'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[2]]).astype(str).str.strip()
            
            # ... (demais mapeamentos que você já tem) ...
            
            # MODULAÇÃO WBC (Coluna 63 / BN)
            df_conferencia['Modulacao WBC'] = df_conferencia[col_boleta].map(df_lookup[df_base.columns[63]]).apply(limpar_modulacao)
            
            # --- CORREÇÃO AQUI: BUSCA PELO ÍNDICE EXATO DAS COLUNAS AC E AD ---
            # AC é a 29ª coluna (índice 28) | AD é a 30ª coluna (índice 29)
            col_min_nome = df_base.columns[28]
            col_max_nome = df_base.columns[29]
            
            df_conferencia['Modulação Mínima'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_min_nome]), errors='coerce').fillna(0)
            df_conferencia['Modulação Máxima'] = pd.to_numeric(df_conferencia[col_boleta].map(df_lookup[col_max_nome]), errors='coerce').fillna(0)
            # -----------------------------------------------------------------

            # ... (restante do código de Cliq e Pendências) ...

            # --- DEFINIÇÃO DA ORDEM (MUITO IMPORTANTE) ---
            # Verifique se todos os nomes aqui batem EXATAMENTE com as chaves criadas acima
            ordem = [
                col_boleta, 'Operacao', 'Tipo Energia', 'Parte', 'Contraparte', 
                'CP/LP', 'CNPJ Contraparte', 'Submercado', 'Montante MWh', 'Volume MWm', 
                'CliqCCEE Paradigma', 'Modulacao WBC', 
                'Modulação Mínima', 'Modulação Máxima', # <--- Devem estar aqui
                'Vendedor', 'Comprador', 'Contrato CliqCCEE', 'Situacao ERP', 
                'Razao Social', 'Pendência Financeira'
            ]
            
            # Filtrar apenas as colunas que realmente existem para evitar erro de visualização
            colunas_visiveis = [c for c in ordem if c in df_conferencia.columns]

            st.dataframe(
                df_conferencia[colunas_visiveis].sort_values(by=col_boleta), 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "Montante MWh": st.column_config.NumberColumn(format="%.3f"), 
                    "Volume MWm": st.column_config.NumberColumn(format="%.6f"),
                    "Modulação Mínima": st.column_config.NumberColumn(format="%.2f"),
                    "Modulação Máxima": st.column_config.NumberColumn(format="%.2f"),
                    "Pendência Financeira": st.column_config.NumberColumn(format="R$ %.2f")
                }
            )
