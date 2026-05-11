def render_boletas_divergentes(titulo, perfis, df_wbc_base, db_keys, filtro_sub="Todos"):
    """
    Exibe boletas que contribuem para NET ≠ 0 em cada perfil.
    Compara Volume MWm da boleta com o esperado e sinaliza o que verificar.
    """
    with st.expander(f"📋 Boletas para Verificar — {titulo}", expanded=False):
        resultados = []
        for perfil in perfis:
            # NET CCEE esperado para este perfil
            comp_ccee = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_COMPRADOR', filtro_sub, db_keys)
            vend_ccee = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_VENDEDOR',  filtro_sub, db_keys)
            net_ccee  = round(comp_ccee - vend_ccee, 6)

            # Boletas WBC deste perfil (comprador ou vendedor)
            mask_c = df_wbc_base['Comprador'].astype(str).str.strip().str.upper() == perfil.upper()
            mask_v = df_wbc_base['Vendedor'].astype(str).str.strip().str.upper()  == perfil.upper()
            df_perfil = df_wbc_base[mask_c | mask_v].copy()

            if df_perfil.empty and abs(net_ccee) < 1e-5:
                continue

            col_bol = df_perfil.columns[0] if not df_perfil.empty else None

            for _, row in df_perfil.iterrows():
                vol   = pd.to_numeric(row.get('Volume MWm', 0), errors='coerce') or 0.0
                lado  = 'Comprador' if str(row.get('Comprador','')).strip().upper() == perfil.upper() else 'Vendedor'
                cliq  = row.get('Contrato CliqCCEE', '-')
                val_v = row.get('Validação Volume', '-')
                resultados.append({
                    'Perfil':             perfil,
                    'Boleta':             row.iloc[0] if col_bol else '-',
                    'Operacao':           row.get('Operacao', '-'),
                    'Lado':               lado,
                    'Contraparte':        row.get('Contraparte', '-'),
                    'Submercado':         row.get('Submercado', '-'),
                    'Volume MWm (WBC)':   round(vol, 6),
                    'Contrato CliqCCEE':  cliq,
                    'Validação Volume':   val_v,
                    'NET CCEE Esperado':  net_ccee,
                    'Status':             '⚠️ Verificar' if (cliq in ['Verificar', '-', ''] or val_v == 'VERIFICAR') else '✅ OK'
                })

            # Se não há boletas mas CCEE tem volume → aponta como pendência
            if df_perfil.empty and abs(net_ccee) > 1e-5:
                resultados.append({
                    'Perfil':             perfil,
                    'Boleta':             '—',
                    'Operacao':           '—',
                    'Lado':               '—',
                    'Contraparte':        '—',
                    'Submercado':         '—',
                    'Volume MWm (WBC)':   0.0,
                    'Contrato CliqCCEE':  '—',
                    'Validação Volume':   '—',
                    'NET CCEE Esperado':  net_ccee,
                    'Status':             '⚠️ Sem boleta WBC'
                })

        if not resultados:
            st.success("Nenhuma boleta para verificar.")
            return

        df_res = pd.DataFrame(resultados)

        # Destaca apenas as que precisam de atenção
        df_verificar = df_res[df_res['Status'] != '✅ OK']
        df_ok        = df_res[df_res['Status'] == '✅ OK']

        if df_verificar.empty:
            st.success("Todas as boletas estão OK.")
        else:
            st.warning(f"{len(df_verificar)} boleta(s) requerem verificação.")
            st.dataframe(
                df_verificar.reset_index(drop=True),
                use_container_width=True, hide_index=True,
                column_config={
                    'Volume MWm (WBC)':  st.column_config.NumberColumn(format="%.6f"),
                    'NET CCEE Esperado': st.column_config.NumberColumn(format="%.6f"),
                }
            )

        if not df_ok.empty:
            with st.expander("Ver boletas OK", expanded=False):
                st.dataframe(df_ok.reset_index(drop=True), use_container_width=True, hide_index=True,
                             column_config={
                                 'Volume MWm (WBC)':  st.column_config.NumberColumn(format="%.6f"),
                                 'NET CCEE Esperado': st.column_config.NumberColumn(format="%.6f"),
                             })


# ---------------------------------------------------------------------------
# MATCH DE CONTRATOS DE COMPRA
# ---------------------------------------------------------------------------

def _match_boletas_cliqs(boletas, cliqs, tem_volume_ccee):
    """
    Algoritmo de match entre boletas (WBC) e contratos CLIQ.

    Parâmetros
    ----------
    boletas : list[dict]
        Cada item deve conter ao menos:
            'id'        – identificador único da boleta
            'volume'    – float com o volume MWm da boleta
    cliqs : list[dict]
        Cada item deve conter ao menos:
            'id'        – identificador único do CLIQ
            'volume'    – float com o volume MWm do contrato
    tem_volume_ccee : bool
        True  → o volume CCEE já está preenchido; o match deve respeitar volumes.
        False → volume CCEE zerado; prioridade é garantir que toda boleta tenha
                pelo menos um CLIQ vinculado (volumes são ignorados).

    Retorna
    -------
    list[dict] com as linhas do match:
        'boleta_id', 'cliq_id', 'volume_match', 'status'
    """
    resultado = []

    # ------------------------------------------------------------------ #
    # CASO 1 – Volume CCEE zerado: vincular cada boleta a ≥1 CLIQ,        #
    #          sem restrição de volume. CLIQs podem ser reutilizados.      #
    # ------------------------------------------------------------------ #
    if not tem_volume_ccee:
        if not cliqs:
            for b in boletas:
                resultado.append({
                    'boleta_id':    b['id'],
                    'cliq_id':      '—',
                    'volume_match': 0.0,
                    'status':       '⚠️ Sem CLIQ disponível',
                })
            return resultado

        cliq_idx = 0
        for b in boletas:
            # Distribui CLIQs sequencialmente; reutiliza quando esgotam
            cliq = cliqs[cliq_idx % len(cliqs)]
            resultado.append({
                'boleta_id':    b['id'],
                'cliq_id':      cliq['id'],
                'volume_match': 0.0,               # volume irrelevante neste caso
                'status':       '✅ Vinculado (sem volume CCEE)',
            })
            # Avança para o próximo CLIQ apenas se ainda houver disponíveis
            if cliq_idx < len(cliqs) - 1:
                cliq_idx += 1
        return resultado

    # ------------------------------------------------------------------ #
    # CASO 2 – Volume CCEE preenchido: match volumétrico                   #
    #                                                                      #
    # Estratégia:                                                          #
    #   • Tenta cobrir cada boleta com um único CLIQ de volume exato.      #
    #   • Se não encontrar, compõe com múltiplos CLIQs (N→1).              #
    #   • Permite também que múltiplas boletas compartilhem um CLIQ (1→N). #
    #   • CLIQs são reutilizados apenas quando não há contratos suficientes.#
    # ------------------------------------------------------------------ #

    # Controle de volume restante por CLIQ (cópia para não mutar o original)
    cliq_restante = {c['id']: c['volume'] for c in cliqs}
    cliq_lista    = [c['id'] for c in cliqs]   # ordem original

    def _pode_reutilizar():
        """True quando há menos CLIQs únicos do que boletas."""
        return len(cliqs) < len(boletas)

    def _repor_cliq(cid):
        """Reabastece o volume de um CLIQ para reutilização."""
        vol_original = next((c['volume'] for c in cliqs if c['id'] == cid), 0.0)
        cliq_restante[cid] = vol_original

    for b in boletas:
        vol_boleta  = round(b['volume'], 6)
        vol_restante_boleta = vol_boleta

        if vol_boleta <= 0:
            resultado.append({
                'boleta_id':    b['id'],
                'cliq_id':      '—',
                'volume_match': 0.0,
                'status':       '⚠️ Volume boleta zerado/inválido',
            })
            continue

        # --- Tentativa 1: um único CLIQ cobre o volume exato da boleta --- #
        cliq_exato = next(
            (cid for cid in cliq_lista
             if abs(cliq_restante.get(cid, 0.0) - vol_boleta) < 1e-5),
            None
        )
        if cliq_exato is not None:
            cliq_restante[cliq_exato] -= vol_boleta
            resultado.append({
                'boleta_id':    b['id'],
                'cliq_id':      cliq_exato,
                'volume_match': vol_boleta,
                'status':       '✅ Match exato',
            })
            continue

        # --- Tentativa 2: composição de múltiplos CLIQs (N→1 boleta) ---- #
        linhas_parciais = []
        cliqs_usados    = []

        for cid in cliq_lista:
            if vol_restante_boleta <= 1e-6:
                break
            disp = cliq_restante.get(cid, 0.0)
            if disp <= 1e-6:
                # CLIQ esgotado; reutiliza se permitido
                if _pode_reutilizar():
                    _repor_cliq(cid)
                    disp = cliq_restante[cid]
                else:
                    continue

            alocado = round(min(disp, vol_restante_boleta), 6)
            cliq_restante[cid]  -= alocado
            vol_restante_boleta  = round(vol_restante_boleta - alocado, 6)

            linhas_parciais.append({
                'boleta_id':    b['id'],
                'cliq_id':      cid,
                'volume_match': alocado,
                'status':       '✅ Match parcial composto',
            })
            cliqs_usados.append(cid)

        if abs(vol_restante_boleta) < 1e-5:
            resultado.extend(linhas_parciais)
        else:
            # Não foi possível cobrir o volume total
            resultado.extend(linhas_parciais)
            resultado.append({
                'boleta_id':    b['id'],
                'cliq_id':      '—',
                'volume_match': round(vol_boleta - vol_restante_boleta, 6),
                'status':       f'⚠️ Volume descoberto: {round(vol_restante_boleta, 6)} MWm',
            })

    return resultado


def render_match_contratos_compra(titulo, perfis, df_wbc_base, df_cliqs, db_keys, filtro_sub="Todos"):
    """
    Exibe a área de match entre boletas de compra (WBC) e contratos CLIQ.

    Parâmetros
    ----------
    titulo      : str          – título do expander
    perfis      : list[str]    – lista de perfis a processar
    df_wbc_base : pd.DataFrame – boletas WBC (mesma base usada em render_boletas_divergentes)
    df_cliqs    : pd.DataFrame – contratos CLIQ disponíveis.
                                 Colunas esperadas: 'Contrato CliqCCEE', 'Volume MWm', 'Perfil'
                                 (ajuste os nomes conforme seu DataFrame real)
    db_keys     : dict         – chaves de acesso ao banco (repassado para calcular_mwmedio_em_bases)
    filtro_sub  : str          – filtro de submercado (default "Todos")
    """
    with st.expander(f"🔗 Match Contratos de Compra — {titulo}", expanded=False):

        todos_matches = []

        for perfil in perfis:
            # ---- Volume CCEE para decidir modo do match ---- #
            comp_ccee = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_COMPRADOR', filtro_sub, db_keys)
            vend_ccee = calcular_mwmedio_em_bases(perfil, 'SIGLA_PERFIL_VENDEDOR',  filtro_sub, db_keys)
            net_ccee  = round(comp_ccee - vend_ccee, 6)
            tem_volume_ccee = abs(net_ccee) > 1e-5

            # ---- Boletas de compra deste perfil ---- #
            mask_c    = df_wbc_base['Comprador'].astype(str).str.strip().str.upper() == perfil.upper()
            df_compra = df_wbc_base[mask_c].copy()

            boletas = [
                {
                    'id':     str(row.iloc[0]),
                    'volume': pd.to_numeric(row.get('Volume MWm', 0), errors='coerce') or 0.0,
                }
                for _, row in df_compra.iterrows()
            ]

            # ---- CLIQs disponíveis para este perfil ---- #
            # Ajuste o filtro de perfil conforme a coluna real do seu df_cliqs
            col_perfil_cliq = 'Perfil' if 'Perfil' in df_cliqs.columns else df_cliqs.columns[0]
            mask_cliq = df_cliqs[col_perfil_cliq].astype(str).str.strip().str.upper() == perfil.upper()
            df_cliq_perfil = df_cliqs[mask_cliq].copy()

            cliqs = [
                {
                    'id':     str(row.get('Contrato CliqCCEE', f'CLIQ_{i}')),
                    'volume': pd.to_numeric(row.get('Volume MWm', 0), errors='coerce') or 0.0,
                }
                for i, (_, row) in enumerate(df_cliq_perfil.iterrows())
            ]

            if not boletas:
                continue

            # ---- Executa o match ---- #
            matches = _match_boletas_cliqs(boletas, cliqs, tem_volume_ccee)

            for m in matches:
                m['Perfil']          = perfil
                m['NET CCEE']        = net_ccee
                m['Modo Match']      = 'Com Volume CCEE' if tem_volume_ccee else 'Sem Volume CCEE'
                todos_matches.append(m)

        # ---- Exibição ---- #
        if not todos_matches:
            st.info("Nenhuma boleta de compra encontrada para os perfis selecionados.")
            return

        df_match = pd.DataFrame(todos_matches).rename(columns={
            'boleta_id':    'Boleta',
            'cliq_id':      'CLIQ Vinculado',
            'volume_match': 'Volume Match (MWm)',
            'status':       'Status Match',
        })

        # Colunas finais em ordem legível
        cols_order = ['Perfil', 'Boleta', 'CLIQ Vinculado', 'Volume Match (MWm)',
                      'NET CCEE', 'Modo Match', 'Status Match']
        df_match = df_match[[c for c in cols_order if c in df_match.columns]]

        # Separar por status para destaque visual
        df_ok_m   = df_match[df_match['Status Match'].str.startswith('✅')]
        df_warn_m = df_match[~df_match['Status Match'].str.startswith('✅')]

        col_cfg = {
            'Volume Match (MWm)': st.column_config.NumberColumn(format="%.6f"),
            'NET CCEE':           st.column_config.NumberColumn(format="%.6f"),
        }

        if not df_warn_m.empty:
            st.warning(f"{len(df_warn_m)} vínculo(s) com pendência.")
            st.dataframe(df_warn_m.reset_index(drop=True),
                         use_container_width=True, hide_index=True,
                         column_config=col_cfg)

        if not df_ok_m.empty:
            lbl = "Ver vínculos OK" if not df_warn_m.empty else "Vínculos de Match"
            expanded_ok = df_warn_m.empty   # abre direto se não há pendências
            with st.expander(lbl, expanded=expanded_ok):
                st.dataframe(df_ok_m.reset_index(drop=True),
                             use_container_width=True, hide_index=True,
                             column_config=col_cfg)

        # ---- Resumo por perfil ---- #
        with st.expander("📊 Resumo do Match por Perfil", expanded=False):
            resumo = (
                df_match
                .groupby(['Perfil', 'Modo Match'])
                .agg(
                    Boletas=('Boleta', 'nunique'),
                    CLIQs_Vinculados=('CLIQ Vinculado', 'nunique'),
                    Volume_Total_Match=('Volume Match (MWm)', 'sum'),
                    Pendencias=('Status Match', lambda x: (x.str.startswith('⚠️')).sum()),
                )
                .reset_index()
            )
            st.dataframe(resumo, use_container_width=True, hide_index=True,
                         column_config={
                             'Volume_Total_Match': st.column_config.NumberColumn(format="%.6f"),
                         })
