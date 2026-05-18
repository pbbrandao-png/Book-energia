# =============================================================================
# MATCH CLIQ BISMUT
# =============================================================================

if df_bismut is not None:

    try:

        # ==========================================
        # LIMPA NOMES DAS COLUNAS
        # ==========================================

        df_bismut.columns = [
            limpar_coluna(col)
            for col in df_bismut.columns
        ]

        # ==========================================
        # FILTRA APENAS BISMUT
        # ==========================================

        df_bismut = df_bismut[
            df_bismut['PARTE']
            .astype(str)
            .str.upper()
            == 'BISMUT COMERCIALIZADORA DE ENERGIA S/A'
        ]

        # ==========================================
        # LIMPA COLUNAS PRINCIPAIS
        # ==========================================

        df['CLIQ PARADIGMA'] = (
            df['CLIQ PARADIGMA']
            .astype(str)
            .str.strip()
        )

        df['Cliq Mês Anterior'] = (
            df['Cliq Mês Anterior']
            .astype(str)
            .str.strip()
        )

        df['VENDEDOR'] = (
            df['VENDEDOR']
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df['COMPRADOR'] = (
            df['COMPRADOR']
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df_bismut['CODIGO_CONTRATO'] = (
            df_bismut['CODIGO_CONTRATO']
            .astype(str)
            .str.strip()
        )

        df_bismut['SIGLA_PERFIL_VENDEDOR'] = (
            df_bismut['SIGLA_PERFIL_VENDEDOR']
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df_bismut['SIGLA_PERFIL_COMPRADOR'] = (
            df_bismut['SIGLA_PERFIL_COMPRADOR']
            .astype(str)
            .str.strip()
            .str.upper()
        )

        # ==========================================
        # FUNÇÃO MATCH
        # ==========================================

        def localizar_cliq(linha):

            cliq_atual = linha['CLIQ PARADIGMA']
            cliq_anterior = linha['Cliq Mês Anterior']
            vendedor = linha['VENDEDOR']
            comprador = linha['COMPRADOR']

            # ======================================
            # PRIORIDADE 1 -> CLIQ PARADIGMA
            # ======================================

            if cliq_atual not in ['', '-', 'nan', 'None']:

                resultado = df_bismut[
                    (
                        df_bismut['CODIGO_CONTRATO']
                        == cliq_atual
                    )
                    &
                    (
                        df_bismut['SIGLA_PERFIL_VENDEDOR']
                        == vendedor
                    )
                    &
                    (
                        df_bismut['SIGLA_PERFIL_COMPRADOR']
                        == comprador
                    )
                ]

                if not resultado.empty:
                    return cliq_atual

            # ======================================
            # PRIORIDADE 2 -> CLIQ MÊS ANTERIOR
            # ======================================

            if cliq_anterior not in ['', '-', 'nan', 'None']:

                resultado = df_bismut[
                    (
                        df_bismut['CODIGO_CONTRATO']
                        == cliq_anterior
                    )
                    &
                    (
                        df_bismut['SIGLA_PERFIL_VENDEDOR']
                        == vendedor
                    )
                    &
                    (
                        df_bismut['SIGLA_PERFIL_COMPRADOR']
                        == comprador
                    )
                ]

                if not resultado.empty:
                    return cliq_anterior

            # ======================================
            # NÃO ENCONTROU
            # ======================================

            return 'VERIFICAR'

        # ==========================================
        # CRIA COLUNA FINAL
        # ==========================================

        df['CLIQ CCEE'] = df.apply(
            localizar_cliq,
            axis=1
        )

        st.success('✅ Match Bismut realizado!')

    except Exception as erro:

        st.warning(
            f'⚠️ Erro no match Bismut: {erro}'
        )

