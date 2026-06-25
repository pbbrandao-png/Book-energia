# APP_BOOK_ENERGIA_V20 - VERSÃO DE ALTA PERFORMANCE (OTIMIZADA)
# Coluna "Contrato CliqCCEE" via CSVs extraídos dos ZIPs Matrix e Bismut
# Boletas ACR (lista fixa) → ccear_q (extraído do ZIP Matrix)
# Matrix (não-Bismut, não-ACR) → cceal_firme + cbr_mercado_proprio (ZIP Matrix)
# Bismut → cceal_firme (ZIP Bismut)
# V17: + Contraparte Razão Social | highlight amarelo Parte==Contraparte | flag ocultar zerados
# V19: + Ajuste Manual via Planilha | Separação de Erros de Divergência e Sem Match Nenhum nos CSVs
# V20: + Otimização massiva de performance + Regra de ignorar Intraportfólio/Zerados nas tabelas de erro

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
    """Lê bytes de um CSV CCEE e limpa colunas."""
    df = pd.read_csv(BytesIO(bytes_csv), sep='\t', encoding='latin1', skiprows=1, dtype=str)
    df.columns = df.columns.str.strip()
    
    # Filtrar rascunhos logo na leitura economiza muita memória e processamento
    if 'SITUACAO_CONTRATO' in df.columns:
        df = df[df['SITUACAO_CONTRATO'].str.strip().str.lower() != 'rascunho']
        
    for col in ['CODIGO_CONTRATO', 'SIGLA_PERFIL_VENDEDOR', 'SIGLA_PERFIL_COMPRADOR', 'SUBMERCADO_ENTREGA', 'MWmedio',
                'LIMITE_MINIMO_MODULACAO_MW', 'LIMITE_MAXIMO_MODULACAO_MW', 'TIPO_MODULACAO']:
        if col in df.columns:
            df[col] = df[col].str.strip()
            
    df['_CHAVE'] = (
        df['SIGLA_PERFIL_VENDEDOR'].fillna('')
        + df['SIGLA_PERFIL_COMPRADOR'].fillna('')
        + df['SUBMERCADO_ENTREGA'].fillna('')
    )
    return df


def extrair_csvs_zip(zip_file):
    result = {'cceal': None, 'cbr': None, 'ccear_q': None}
    if zip_file is None:
        return result
    try:
        with zipfile.ZipFile(zip_file) as zf:
            for nome in zf.namelist():
                nome_lower = nome.lower()
                if nome_lower.endswith('/') or not nome_lower.endswith('.csv') or 'parcela' in nome_lower:
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
    validos = [df for df in lista if df is not None and not df.empty]
    if not validos:
        return pd.DataFrame()
    return pd.concat(validos, ignore_index=True)


def criar_indices_busca(df_ccee):
    """Mapeia os códigos da CCEE em dicionários para busca em tempo de execução O(1)."""
    if df_ccee.empty:
        return {}, {}, {}, {}, {}
    
    # Remove duplicados mantendo o primeiro registro válido
    df_limpo = df_ccee.drop_duplicates(subset=['CODIGO_CONTRATO'])
    
    dict_chave = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo['_CHAVE']))
    dict_vend = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_VENDEDOR', '')))
    dict_comp = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SIGLA_PERFIL_COMPRADOR', '')))
    dict_sub = dict(zip(df_limpo['CODIGO_CONTRATO'], df_limpo.get('SUBMERCADO_ENTREGA', '')))
    
    # Conjunto para checar existência imediata
    set_existentes = set(df_limpo['CODIGO_CONTRATO'])
    
    return dict_chave, dict_vend, dict_comp, dict_sub, set_existentes


# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Book Energia", layout="wide")

pagina = st.sidebar.radio("Menu", ["Base Conferência", "Encontro Energético", "Arquivos CCEE"])
st.title("📊 Book Energia")

arquivo = st.file_uploader("Selecione a RelPers", type=["xlsx", "xlsm"])
arquivo_mes_anterior = st.file_uploader("Selecione a planilha Mês Anterior", type=["xlsx"])
arquivo_ajuste_manual = st.file_uploader("Selecione a planilha de Ajuste Manual", type=["xlsx"])
zip_matrix = st.file_uploader("Selecione o ZIP Matrix", type=["zip"])
zip_bismut = st.file_uploader("Selecione o ZIP Bismut", type=["zip"])

if arquivo is not None:
    try:
        df = pd.read_excel(arquivo, header=8)

        horas_mes = {
            1: 744, 2: 672, 3: 744, 4: 720, 5: 744, 6: 720,
            7: 744, 8: 744, 9: 720, 10: 744, 11: 720, 12: 744
        }

        if arquivo_mes_anterior is not None:
            df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
            mapa_mes_anterior = dict(zip(df_mes_anterior["BOLETA"], df_mes_anterior["Codigo_CCEE"]))
        else:
            mapa_mes_anterior = {}

        mapa_ajuste_manual_paradigma = {}
        mapa_ajuste_manual_contraparte = {}
        if arquivo_ajuste_manual is not None:
            try:
                df_aj_manual = pd.read_excel(arquivo_ajuste_manual)
                if "BOLETA" in df_aj_manual.columns:
                    df_aj_manual = df_aj_manual.dropna(subset=["BOLETA"])
                    df_aj_manual["BOLETA"] = df_aj_manual["BOLETA"].astype(str).str.strip().str.replace(".0", "", regex=False)
                    if "CliqCCEE Paradigma" in df_aj_manual.columns:
                        for _, row in df_aj_manual.iterrows():
                            mapa_ajuste_manual_paradigma[str(row["BOLETA"])] = str(row["CliqCCEE Paradigma"])
                    if "Contraparte" in df_aj_manual.columns:
                        for _, row in df_aj_manual.iterrows():
                            mapa_ajuste_manual_contraparte[str(row["BOLETA"])] = str(row["Contraparte"])
            except Exception as e:
                st.warning(f"Erro ao processar ajustes manuais: {e}")

        csvs = {}
        if zip_matrix is not None:
            csvs_matrix = extrair_csvs_zip(zip_matrix)
            csvs.update(csvs_matrix)
        if zip_bismut is not None:
            csvs_bismut = extrair_csvs_zip(zip_bismut)
            if csvs_bismut['cceal']:
                csvs['cceal_bismut'] = csvs_bismut['cceal']

        idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext = criar_indices_busca(csvs.get('ccear_q'))
        idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext = criar_indices_busca(combiner_dfs([csvs.get('cceal'), csvs.get('cbr')]))
        idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext = criar_indices_busca(csvs.get('cceal_bismut'))
        csvs_disponiveis = any([csvs.get('ccear_q'), csvs.get('cceal'), csvs.get('cbr'), csvs.get('cceal_bismut')])

        def calcular_contrato_cliqccee_fast(row):
            boleta = int(float(str(row["BOLETA"]).strip())) if row["BOLETA"] else -1
            chave = ""
            if boleta in BOLETAS_ACR: chave = idx_a_chave.get(str(boleta), "")
            elif str(row.get("Parte", "")).strip().upper() == "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A.": chave = idx_b_chave.get(str(boleta), "")
            else: chave = idx_m_chave.get(str(boleta), "")
            return f"{boleta}{chave}" if boleta > 0 and chave else ""

        nets = df.dropna(subset=["Parte", "Contraparte"])
        base = df.copy()
        base["Contrato CliqCCEE"] = base.apply(calcular_contrato_cliqccee_fast, axis=1)
        base["Contrato CliqCCEE"] = base.apply(lambda row: mapa_mes_anterior.get(str(row["BOLETA"]), row["Contrato CliqCCEE"]) if row["Contrato CliqCCEE"] == "" else row["Contrato CliqCCEE"], axis=1)
        base["CliqCCEE Paradigma"] = base["BOLETA"].astype(str).map(mapa_ajuste_manual_paradigma).fillna(base.get("CliqCCEE Paradigma", ""))
        base["Contraparte"] = base["BOLETA"].astype(str).map(mapa_ajuste_manual_contraparte).fillna(base.get("Contraparte", ""))
        base["Editado Manualmente"] = False
        if "Base Conferência" in st.session_state:
            base = st.session_state["Base Conferência"]
        if "base_editada" in st.session_state:
            base = st.session_state["base_editada"]

        if pagina == "Base Conferência":
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                opcoes_filtro_parte = ["Todas"] + sorted(base["Parte"].dropna().unique().tolist())
                filtro_parte = st.selectbox("Filtrar por Parte", opcoes_filtro_parte, label_visibility="collapsed")
            with col2:
                opcoes_filtro_operacao = ["Todas"] + sorted(base["Operação"].dropna().unique().tolist())
                filtro_operacao = st.selectbox("Filtrar por Operação", opcoes_filtro_operacao, label_visibility="collapsed")
            with col3:
                flag_ocultar_zerados = st.checkbox("Ocultar zerados")

            base_filtrada = base if filtro_parte == "Todas" else base[base["Parte"] == filtro_parte]
            base_filtrada = base_filtrada if filtro_operacao == "Todas" else base_filtrada[base_filtrada["Operação"] == filtro_operacao]
            if flag_ocultar_zerados:
                base_filtrada = base_filtrada[base_filtrada["Volume (MWh)"] != 0.0]

            st.caption(f"📊 {len(base_filtrada):,} contrato(s) carregado(s)")

            flag_mesmo_titular = st.checkbox("Destacar Parte == Contraparte (amarelo)")

            base_exibicao = base_filtrada[["BOLETA", "Operação", "Tipo de Energia", "Parte", "Contraparte Razão Social", "Contraparte", "CP/LP", "CNPJ CONTRAPARTE", "Submercado", "Volume (MWh)", "Volume MWm", "CliqCCEE Paradigma", "Modulação WBC", "% Modulação Mínima", "% Modulação Máxima", "Contrato CliqCCEE mês anterior", "Vendedor", "Comprador", "Contrato CliqCCEE", "Editado Manualmente", "Volume Book", "Volume CCEE", "Check Volume", "Volume Global", "Volume Global CCEE", "Check Volume Global", "Modulação Mínima", "Modulação Máxima", "Modulação Mínima CCEE", "Modulação Máxima CCEE", "Modulação CCEE", "Check Modulação", "Check Modulação Mínima", "Check Modulação Máxima"]].reset_index(drop=True)

            col_config = {
                "BOLETA": st.column_config.Column(disabled=True), "Operação": st.column_config.Column(disabled=True),
                "Tipo de Energia": st.column_config.Column(disabled=True), "Parte": st.column_config.Column(disabled=True),
                "Contraparte Razão Social": st.column_config.Column(disabled=True), "Contraparte": st.column_config.TextColumn(disabled=False),
                "CP/LP": st.column_config.Column(disabled=True), "CNPJ CONTRAPARTE": st.column_config.Column(disabled=True),
                "Submercado": st.column_config.Column(disabled=True), "Volume (MWh)": st.column_config.Column(disabled=True),
                "Volume MWm": st.column_config.Column(disabled=True), "CliqCCEE Paradigma": st.column_config.TextColumn(disabled=False),
                "Modulação WBC": st.column_config.Column(disabled=True), "% Modulação Mínima": st.column_config.Column(disabled=True),
                "% Modulação Máxima": st.column_config.Column(disabled=True), "Contrato CliqCCEE mês anterior": st.column_config.TextColumn(disabled=True),
                "Vendedor": st.column_config.TextColumn(disabled=True), "Comprador": st.column_config.TextColumn(disabled=True),
                "Contrato CliqCCEE": st.column_config.TextColumn(disabled=True), "Editado Manualmente": st.column_config.Column(disabled=True),
                "Volume Book": st.column_config.Column(disabled=True), "Volume CCEE": st.column_config.Column(disabled=True),
                "Check Volume": st.column_config.Column(disabled=True), "Volume Global": st.column_config.Column(disabled=True),
                "Volume Global CCEE": st.column_config.Column(disabled=True), "Check Volume Global": st.column_config.Column(disabled=True),
                "Modulação Mínima": st.column_config.Column(disabled=True), "Modulação Máxima": st.column_config.Column(disabled=True),
                "Modulação Mínima CCEE": st.column_config.Column(disabled=True), "Modulação Máxima CCEE": st.column_config.Column(disabled=True),
                "Modulação CCEE": st.column_config.Column(disabled=True), "Check Modulação": st.column_config.Column(disabled=True),
                "Check Modulação Mínima": st.column_config.Column(disabled=True), "Check Modulação Máxima": st.column_config.Column(disabled=True),
            }

            # ✅ CORRIGIDO: Remove styling complexo, passa só os dados
            base_editada_df = st.data_editor(base_exibicao, use_container_width=True, hide_index=True, column_config=col_config, key="editor_base")

            if st.session_state.get("editor_base") and st.session_state["editor_base"].get("edited_rows"):
                edicoes = st.session_state["editor_base"]["edited_rows"]
                indices_exibicao = base_exibicao.index.tolist()
                if "contratos_editados_diretamente" not in st.session_state:
                    st.session_state["contratos_editados_diretamente"] = []

                for idx_str, alteracoes in edicoes.items():
                    idx = int(idx_str)
                    idx_real = indices_exibicao[idx]
                    boleta_alvo = base.loc[idx_real, "BOLETA"]
                    base.loc[idx_real, "Editado Manualmente"] = True
                    for col, val in alteracoes.items():
                        base.loc[idx_real, col] = str(val)
                        if col == "Contrato CliqCCEE" and boleta_alvo not in st.session_state["contratos_editados_diretamente"]:
                            st.session_state["contratos_editados_diretamente"].append(boleta_alvo)
                    if csvs_disponiveis and "Contrato CliqCCEE" not in alteracoes:
                        if boleta_alvo not in st.session_state["contratos_editados_diretamente"]:
                            base.loc[idx_real, "Contrato CliqCCEE"] = str(calcular_contrato_cliqccee_fast(base.loc[idx_real]))
                st.session_state["base_editada"] = base.copy()
                st.rerun()

            base_download = base.copy()
            if flag_ocultar_zerados: base_download = base_download[base_download["Volume (MWh)"] != 0.0]
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer: base_download.to_excel(writer, sheet_name="Base Conferência", index=False)
            st.download_button("📥 Download Base Conferência", data=output.getvalue(), file_name="Base_Conferencia.xlsx")

            # ── CLASSIFICAÇÃO ULTRA RÁPIDA DE ERROS / SEM MATCH (Mapeamento Vetorizado) ──
            if csvs_disponiveis:
                st.markdown("---")
                lista_divergencias = []
                lista_sem_match_nenhum = []

                BISMUT_NOME_UPPER = "NEWAVE BISMUT COMERCIALIZADORA DE ENERGIA S.A."

                # Varre a lista processando de maneira imediata com os dicionários em memória
                for _, row in base.iterrows():
                    parte_limpa = str(row["Parte"]).strip().upper()
                    contraparte_limpa = str(row["Contraparte Razão Social"]).strip().upper()
                    
                    # Regra de ouro: Se for Intraportfólio (titular igual) ou o volume for 0, não precisa de match!
                    if parte_limpa == contraparte_limpa or float(row["Volume (MWh)"]) == 0.0:
                        continue

                    try: b_int = int(float(str(row["BOLETA"]).strip()))
                    except: b_int = -1

                    if b_int in BOLETAS_ACR:
                        d_ch, d_v, d_c, d_s, s_ext = idx_a_chave, idx_a_v, idx_a_c, idx_a_s, set_a_ext
                    elif parte_limpa == BISMUT_NOME_UPPER:
                        d_ch, d_v, d_c, d_s, s_ext = idx_b_chave, idx_b_v, idx_b_c, idx_b_s, set_b_ext
                    else:
                        d_ch, d_v, d_c, d_s, s_ext = idx_m_chave, idx_m_v, idx_m_c, idx_m_s, set_m_ext

                    cods = [str(row.get(c, "")).strip() for c in ["Contrato CliqCCEE", "Contrato CliqCCEE mês anterior", "CliqCCEE Paradigma"]]
                    cods_validos = [c for c in cods if c not in ('', '-', 'None', 'nan')]

                    if not cods_validos:
                        status, justificativa = "SEM_MATCH", "Contrato inexistente (Sem códigos informados)"
                    else:
                        cod_encontrado = None
                        for c in cods_validos:
                            if c in s_ext:
                                cod_encontrado = c
                                break
                        
                        if not cod_encontrado:
                            status, justificativa = "SEM_MATCH", "Contrato inexistente no CSV CCEE"
                        else:
                            v_b, c_b, s_b = str(row["Vendedor"]).strip(), str(row["Comprador"]).strip(), str(row["Submercado"]).strip()
                            v_c, c_c, s_c = d_v.get(cod_encontrado, ''), d_c.get(cod_encontrado, ''), d_s.get(cod_encontrado, '')

                            divs = []
                            if v_b != v_c: divs.append("Vendedor")
                            if c_b != c_c: divs.append("Comprador")
                            if s_b != s_c: divs.append(f"Submercado (Boleta={s_b} | CSV={s_c})")

                            if not divs:
                                status, justificativa = "OK", None
                            else:
                                status = "ERRO"
                                if len(divs) == 1: justificativa = f"Divergência de {divs[0]}"
                                elif len(divs) == 2: justificativa = f"Divergência de {divs[0]} e {divs[1]}"
                                else: justificativa = f"Divergência de {divs[0]}, {divs[1]} e {divs[2]}"

                    if status in ("ERRO", "SEM_MATCH"):
                        item = {"Boleta": row["BOLETA"], "Vendedor": row["Vendedor"], "Comprador": row["Comprador"], "Mensagem": justificativa}
                        if status == "ERRO": lista_divergencias.append(item)
                        else: lista_sem_match_nenhum.append(item)

                df_divergencias = pd.DataFrame(lista_divergencias, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])
                df_sem_match_nenhum = pd.DataFrame(lista_sem_match_nenhum, columns=["Boleta", "Vendedor", "Comprador", "Mensagem"])

                st.subheader("❌ Contratos com Divergência (Existem no CSV, mas dados não batem)")
                st.caption(f"{len(df_divergencias):,} contrato(s) com divergência encontrado(s)")
                st.dataframe(df_divergencias, use_container_width=True, hide_index=True)

                output_div = BytesIO()
                with pd.ExcelWriter(output_div, engine="openpyxl") as writer: df_divergencias.to_excel(writer, sheet_name="Divergencias", index=False)
                st.download_button("📥 Download Contratos com Divergência", data=output_div.getvalue(), file_name="Contratos_com_Divergencia.xlsx")
                st.markdown("---")

                st.subheader("🔍 Contratos Sem Match Nenhum (Inexistentes no CSV CCEE)")
                st.caption(f"{len(df_sem_match_nenhum):,} contrato(s) não localizados nos arquivos da CCEE")
                st.dataframe(df_sem_match_nenhum, use_container_width=True, hide_index=True)

                output_sm = BytesIO()
                with pd.ExcelWriter(output_sm, engine="openpyxl") as writer: df_sem_match_nenhum.to_excel(writer, sheet_name="Sem Match Nenhum", index=False)
                st.download_button("📥 Download Contratos Sem Match Nenhum", data=output_sm.getvalue(), file_name="Contratos_Sem_Match_Nenhum.xlsx")

        elif pagina == "Encontro Energético":
            st.subheader("🤝 Encontro Energético")
            parte = st.selectbox("Parte", sorted(nets["Parte"].dropna().unique()))
            df_parte = nets[nets["Parte"] == parte]
            contraparte = st.selectbox("Contraparte", sorted(df_parte["Contraparte"].dropna().unique()))
            df_contraparte = df_parte[df_parte["Contraparte"] == contraparte]
            submercado = st.selectbox("Submercado", sorted(df_contraparte["Submercado"].dropna().unique()))
            df_sub = df_contraparte[df_contraparte["Submercado"] == submercado]
            tipo_energia = st.selectbox("Tipo de Energia", sorted(df_sub["Tipo de Energia"].dropna().unique()))

            encontro = base[(base["Parte"] == parte) & (base["Contraparte"] == contraparte) & (base["Submercado"] == submercado) & (base["Tipo de Energia"] == tipo_energia)]
            compras_calc = encontro[encontro["Operação"] == "Compra"]
            vendas_calc  = encontro[encontro["Operação"] == "Venda"]

            compras, vendas = compras_calc.copy(), vendas_calc.copy()
            compras["Volume (MWh)"] = compras["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            compras["Volume MWm"]   = compras["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)
            vendas["Volume (MWh)"]  = vendas["Volume (MWh)"].map(lambda x: f"{x:.3f}" if isinstance(x, (int, float)) else x)
            vendas["Volume MWm"]    = vendas["Volume MWm"].map(lambda x: f"{x:.6f}" if isinstance(x, (int, float)) else x)

            st.markdown("## COMPRAS")
            st.dataframe(compras[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)
            st.markdown("## VENDAS")
            st.dataframe(vendas[["BOLETA", "Volume (MWh)", "Volume MWm"]], hide_index=True, use_container_width=True)

            total_compra, total_venda = compras_calc["Volume (MWh)"].sum(), vendas_calc["Volume (MWh)"].sum()
            saldo = total_compra - total_venda
            total_compra_mwm, total_venda_mwm = compras_calc["Volume MWm"].sum(), vendas_calc["Volume MWm"].sum()
            mes_referencia = int(df["Mes"].dropna().iloc[0])
            saldo_mwm = saldo / horas_mes.get(mes_referencia, 744)

            ajuste = contraparte if saldo > 0 else parte if saldo < 0 else "ZERADO"
            resumo = pd.DataFrame({"Tipo": ["Compras", "Vendas", "Saldo"], "MWh": [f"{total_compra:.3f}", f"{total_venda:.3f}", f"{saldo:.3f}"], "MWm": [f"{total_compra_mwm:.6f}", f"{total_venda_mwm:.6f}", f"{saldo_mwm:.6f}"]})
            st.markdown("## RESUMO")
            st.dataframe(resumo, hide_index=True, use_container_width=True)

            c1, c2 = st.columns(2)
            with c1: st.metric("Quem Ajusta", ajuste)
            with c2: st.metric("Volume a Ajustar (MWm)", f"{abs(saldo_mwm):.6f}")

    except Exception as erro:
        st.error("Erro ao processar a planilha")
        st.exception(erro)
