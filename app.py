# =============================================================================
#
#   BOOK DE ENERGIA
#   ---------------
#   Este arquivo lê planilhas de contratos de energia, trata os dados
#   e exibe o resultado numa tabela dentro do Streamlit.
#
#   ESTRUTURA DO ARQUIVO:
#   1. Importações          → bibliotecas que o código usa
#   2. Funções de limpeza   → deixam os nomes de coluna no padrão
#   3. Funções de conversão → traduzem siglas e calculam valores
#   4. Função principal     → junta tudo e processa a planilha
#   5. Interface            → monta a tela do Streamlit
#
# =============================================================================


# =============================================================================
# 1. IMPORTAÇÕES
#    Cada linha importa uma biblioteca externa que usamos mais abaixo.
# =============================================================================

import streamlit as st   # cria a interface visual (tela, botões, tabelas)
import pandas as pd      # lê e manipula planilhas/tabelas de dados
import unicodedata       # remove acentos de textos
import calendar          # consulta quantos dias tem cada mês


# =============================================================================
# 2. FUNÇÕES DE LIMPEZA DE COLUNAS
#    Usadas para padronizar os nomes das colunas da planilha antes de
#    qualquer processamento. Assim evitamos erros por espaço, acento, etc.
# =============================================================================

def limpar_coluna(texto):
    """
    Recebe o nome de uma coluna e devolve ele padronizado:
      - sem espaços nas bordas
      - em MAIÚSCULAS
      - sem acentos (ex: "Código" → "CODIGO")

    Exemplo de uso:
        limpar_coluna("  fonte_contrato ") → "FONTE_CONTRATO"
    """

    # Remove espaços do início e fim, e deixa tudo maiúsculo
    texto = str(texto).strip().upper()

    # Remove acentos usando o método de normalização Unicode
    texto = (
        unicodedata
        .normalize('NFKD', texto)   # separa letras dos acentos
        .encode('ASCII', 'ignore')  # descarta os acentos
        .decode('utf-8')            # converte de volta para texto normal
    )

    return texto


# =============================================================================
# 3. FUNÇÕES DE CONVERSÃO
#    Cada função recebe um valor "bruto" da planilha e devolve o valor
#    já tratado no padrão do book.
#    Padrão: se o valor não estiver no mapa, devolve ele mesmo sem alterar.
# =============================================================================

# -----------------------------------------------------------------------------
# 3a. FONTE
#     Converte o nome longo da fonte para a sigla usada no book.
# -----------------------------------------------------------------------------

def tratar_fonte(valor):
    """
    Converte o nome completo da fonte de energia para a sigla padronizada.

    Exemplo:
        "Incentivada 50%"  →  "Incentivada-I5"
        "Incentivada 100%" →  "Incentivada-I1"
    """

    # Dicionário de conversão: chave = valor original, valor = valor desejado
    mapa_fontes = {
        'Incentivada 50%':           'Incentivada-I5',
        'Cogeração Qualificada 50%': 'Incentivada-CQ5',
        'Incentivada 100%':          'Incentivada-I1',
        'Incentivada 0%':            'Incentivada-I0',
    }

    # .get(valor, valor) → busca o valor no mapa;
    # se não encontrar, devolve o valor original sem alterar
    return mapa_fontes.get(valor, valor)


# -----------------------------------------------------------------------------
# 3b. SUBMERCADO
#     Converte a sigla do submercado para o nome por extenso.
# -----------------------------------------------------------------------------

def tratar_submercado(valor):
    """
    Converte a sigla do submercado para o nome completo.

    Exemplo:
        "SE/CO" → "SUDESTE"
        "NE"    → "NORDESTE"
    """

    # Garante que o valor está em maiúsculas e sem espaços extras
    valor = str(valor).strip().upper()

    # Dicionário de conversão: sigla → nome por extenso
    mapa_submercados = {
        'N':     'NORTE',
        'S':     'SUL',
        'NE':    'NORDESTE',
        'SE/CO': 'SUDESTE',
    }

    return mapa_submercados.get(valor, valor)


# -----------------------------------------------------------------------------
# 3c. CP / LP
#     Classifica o contrato como Curto Prazo (CP) ou Longo Prazo (LP)
#     com base na quantidade de dias de suprimento.
# -----------------------------------------------------------------------------

def calcular_cp_lp(dias):
    """
    Recebe a duração do contrato em dias e devolve 'CP' ou 'LP'.

    Regra:
        Até 31 dias  → CP (Curto Prazo)
        Acima de 31  → LP (Longo Prazo)
    """

    # Se o valor estiver vazio/inválido, devolve traço
    if pd.isna(dias):
        return '-'

    # Operador ternário: faz o if/else em uma linha
    # Leitura: "devolve LP se dias > 31, senão devolve CP"
    return 'LP' if dias > 31 else 'CP'


# -----------------------------------------------------------------------------
# 3d. MODULAÇÃO
#     Converte o código de modulação do WBC para o nome padronizado.
# -----------------------------------------------------------------------------

def tratar_modulacao(valor):
    """
    Converte o tipo de modulação do WBC para o padrão do book.

    Exemplo:
        "C - Carga"  → "CARGA"
        "F - Flat"   → "FLAT"
        "DECLARADO"  → "DECLARADA"
    """

    mapa_modulacoes = {
        'C - Carga': 'CARGA',
        'F - Flat':  'FLAT',
        'DECLARADO': 'DECLARADA',
    }

    return mapa_modulacoes.get(valor, valor)


# -----------------------------------------------------------------------------
# 3e. HORAS DO MÊS
#     Calcula quantas horas tem o mês, considerando anos bissextos.
#     Importante para o cálculo de MWm.
# -----------------------------------------------------------------------------

def calcular_horas_do_mes(mes, ano):
    """
    Retorna o total de horas de um mês específico.

    Por que precisamos do ano?
    → Fevereiro tem 28 dias em anos normais e 29 em anos bissextos.
      Usar o ano garante o cálculo correto.

    Exemplo:
        calcular_horas_do_mes(2, 2024) → 696  (fevereiro bissexto: 29 dias × 24h)
        calcular_horas_do_mes(2, 2023) → 672  (fevereiro normal:   28 dias × 24h)
    """

    try:
        mes = int(mes)
        ano = int(ano)

        # calendar.monthrange(ano, mes) devolve uma tupla com
        # (dia da semana do 1º dia, total de dias do mês)
        # O [1] pega só o total de dias
        dias_no_mes = calendar.monthrange(ano, mes)[1]

        return dias_no_mes * 24  # converte dias em horas

    except Exception:
        # Se o mês ou ano for inválido, devolve None (vazio)
        return None


# -----------------------------------------------------------------------------
# 3f. FORMATAÇÃO DE NÚMEROS
#     Formata números no padrão brasileiro: ponto para milhar, vírgula para decimal.
#     ATENÇÃO: use apenas para exibição. Após formatar, o valor vira texto.
# -----------------------------------------------------------------------------

def formatar_numero_br(valor, casas_decimais):
    """
    Formata um número no padrão brasileiro.

    Exemplo:
        formatar_numero_br(1234567.89, 3) → "1.234.567,890"

    Parâmetros:
        valor          → o número a formatar
        casas_decimais → quantas casas depois da vírgula mostrar
    """

    try:
        # Passo 1: formata com separadores no padrão americano (ex: 1,234,567.890)
        # Passo 2: troca vírgula por X temporário (para não confundir na próxima troca)
        # Passo 3: troca ponto por vírgula (padrão BR)
        # Passo 4: troca X de volta para ponto
        return (
            f"{valor:,.{casas_decimais}f}"
            .replace(',', 'X')
            .replace('.', ',')
            .replace('X', '.')
        )

    except Exception:
        # Se o valor não puder ser formatado, devolve ele mesmo
        return valor


# =============================================================================
# 4. FUNÇÃO PRINCIPAL DE PROCESSAMENTO
#    Recebe a planilha bruta e aplica todas as transformações em sequência.
#    Cada bloco dentro desta função cuida de uma etapa específica.
# =============================================================================

def processar_contratos(df):
    """
    Processa o DataFrame bruto de contratos aprovados.

    Etapas:
        1. Copia o DataFrame (protege o original)
        2. Padroniza nomes de colunas
        3. Renomeia colunas para o padrão do book
        4. Trata os valores de cada coluna
        5. Calcula colunas derivadas (dias, CP/LP, horas, MWm)
        6. Cria colunas de exibição formatadas

    Retorna o DataFrame pronto para exibição.
    """

    # Faz uma cópia para não modificar o DataFrame original
    df = df.copy()


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 1 — Padroniza os nomes das colunas
    # Aplica a função limpar_coluna() em cada nome de coluna da planilha.
    # Isso garante que não vamos ter problemas com acentos ou maiúsculas.
    # ─────────────────────────────────────────────────────────────────────

    df.columns = [limpar_coluna(col) for col in df.columns]


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 2 — Renomeia colunas para o padrão do book
    # O dicionário abaixo mapeia: nome original → nome desejado
    # ─────────────────────────────────────────────────────────────────────

    mapa_renomeacao = {
        'PARTE_NOME_FANTASIA':       'PARTE',
        'MOVIMENTACAO':              'OPERACAO',
        'FONTE_CONTRATO':            'FONTE',
        'CODIGO_WBC':                'BOLETA',
        'CONTRAPARTE_NOME_FANTASIA': 'CONTRAPARTE',
        'QUANTATUALIZADA':           'MONTANTE_MWH',
        'CODIGO_CCEE':               'CLIQ PARADIGMA',
        'TIPO_DE_MODULACAO':         'MODULACAO WBC',
        'FLEXLIMITE_MODULACAOMAX':   'MOD MAX',
        'FLEXLIMITE_MODULACAOMIN':   'MOD MIN',
    }

    # Filtra só as colunas que realmente existem na planilha
    # (evita erro silencioso quando uma coluna não é encontrada)
    renomear_valido = {
        nome_original: nome_novo
        for nome_original, nome_novo in mapa_renomeacao.items()
        if nome_original in df.columns
    }

    # Verifica se alguma coluna esperada não foi encontrada e avisa o usuário
    colunas_ausentes = set(mapa_renomeacao.keys()) - set(renomear_valido.keys())
    if colunas_ausentes:
        st.warning(
            f"⚠️ Colunas não encontradas no arquivo: "
            f"{', '.join(sorted(colunas_ausentes))}"
        )

    # Aplica a renomeação
    df = df.rename(columns=renomear_valido)


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 3 — Trata os valores de cada coluna
    # O "if coluna in df.columns" evita erro caso a coluna não exista.
    # O ".apply(função)" aplica a função linha por linha na coluna.
    # ─────────────────────────────────────────────────────────────────────

    if 'FONTE' in df.columns:
        df['FONTE'] = df['FONTE'].apply(tratar_fonte)

    if 'SUBMERCADO' in df.columns:
        df['SUBMERCADO'] = df['SUBMERCADO'].apply(tratar_submercado)

    if 'MODULACAO WBC' in df.columns:
        df['MODULACAO WBC'] = df['MODULACAO WBC'].apply(tratar_modulacao)


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 4 — Calcula DIAS e CP/LP
    # Converte as colunas de data e subtrai uma da outra para obter os dias.
    # ─────────────────────────────────────────────────────────────────────

    colunas_de_data_existem = (
        'SUPRIMENTO_INICIO'  in df.columns
        and
        'SUPRIMENTO_TERMINO' in df.columns
    )

    if colunas_de_data_existem:

        # Converte as colunas para o formato de data (datetime)
        # errors='coerce' → valores inválidos viram NaT (data vazia) em vez de erro
        df['SUPRIMENTO_INICIO'] = pd.to_datetime(df['SUPRIMENTO_INICIO'], errors='coerce')
        df['SUPRIMENTO_TERMINO'] = pd.to_datetime(df['SUPRIMENTO_TERMINO'], errors='coerce')

        # Subtrai as datas para obter a duração em dias
        df['DIAS'] = (df['SUPRIMENTO_TERMINO'] - df['SUPRIMENTO_INICIO']).dt.days

        # Classifica como CP ou LP com base nos dias
        df['CP/LP'] = df['DIAS'].apply(calcular_cp_lp)


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 5 — Calcula HORAS DO MÊS
    # Precisamos do mês e do ano para calcular as horas corretamente.
    # O ano é extraído da coluna SUPRIMENTO_INICIO quando disponível.
    # ─────────────────────────────────────────────────────────────────────

    if 'MES' in df.columns:

        # Garante que MES é numérico (ex: "1", "02" → 1, 2)
        df['MES'] = pd.to_numeric(df['MES'], errors='coerce')

        # Tenta pegar o ano da coluna de início; se não existir, usa o ano atual
        if 'SUPRIMENTO_INICIO' in df.columns:
            df['ANO'] = df['SUPRIMENTO_INICIO'].dt.year
        else:
            df['ANO'] = pd.Timestamp.today().year

        # Aplica a função linha a linha usando os valores de MES e ANO
        # axis=1 significa "percorre linha por linha" (ao invés de coluna por coluna)
        df['HORAS_MES'] = df.apply(
            lambda linha: calcular_horas_do_mes(linha['MES'], linha['ANO']),
            axis=1
        )


    # ─────────────────────────────────────────────────────────────────────
    # ETAPA 6 — Calcula e formata MONTANTE
    # Mantemos colunas numéricas (_NUM) separadas das colunas formatadas.
    # As colunas formatadas são só para exibição — não use em cálculos.
    # ─────────────────────────────────────────────────────────────────────

    if 'MONTANTE_MWH' in df.columns:

        # Cria coluna numérica (para cálculos futuros)
        df['MONTANTE_MWH_NUM'] = pd.to_numeric(df['MONTANTE_MWH'], errors='coerce')

        # Cria coluna formatada para exibição (3 casas decimais, padrão BR)
        df['MONTANTE MWh'] = df['MONTANTE_MWH_NUM'].apply(
            lambda valor: formatar_numero_br(valor, 3)
        )

        # Calcula MWm somente se tivermos as horas do mês
        if 'HORAS_MES' in df.columns:

            # Cria coluna numérica de MWm (para cálculos futuros)
            df['MONTANTE_MWM_NUM'] = df['MONTANTE_MWH_NUM'] / df['HORAS_MES']

            # Cria coluna formatada para exibição (6 casas decimais, padrão BR)
            df['MONTANTE MWm'] = df['MONTANTE_MWM_NUM'].apply(
                lambda valor: formatar_numero_br(valor, 6)
            )

    # Devolve o DataFrame completamente processado
    return df


# =============================================================================
# 5. INTERFACE — STREAMLIT
#    Monta a tela do aplicativo: título, uploads, avisos e tabela de resultados.
# =============================================================================

# Configura a página (deve ser a primeira chamada do Streamlit)
st.set_page_config(
    page_title="Book de Energia",
    layout="wide"   # usa toda a largura da tela
)

# Título principal da página
st.title("⚡ Book de Energia")


# ─────────────────────────────────────────────────────────────────────────────
# UPLOADS
# Exibe os dois botões de upload lado a lado usando colunas do Streamlit.
# ─────────────────────────────────────────────────────────────────────────────

# Divide a tela em 2 colunas de tamanho igual
coluna_esquerda, coluna_direita = st.columns(2)

# Upload do arquivo principal (contratos aprovados)
with coluna_esquerda:
    arquivo_aprovados = st.file_uploader(
        label="Contratos aprovados",
        type=['xlsx', 'csv', 'xlsm'],  # tipos de arquivo aceitos
        key='aprovados'                # identificador único do campo
    )

# Upload do arquivo secundário (mês anterior)
with coluna_direita:
    arquivo_mes_anterior = st.file_uploader(
        label="Contratos mês anterior",
        type=['xlsx', 'csv', 'xlsm'],
        key='mes_anterior'
    )


# ─────────────────────────────────────────────────────────────────────────────
# VALIDAÇÃO DO ARQUIVO PRINCIPAL
# Se o arquivo principal não foi enviado, exibe uma mensagem e para aqui.
# O st.stop() impede que o restante do código rode sem o arquivo.
# ─────────────────────────────────────────────────────────────────────────────

if arquivo_aprovados is None:
    st.info("📂 Faça o upload do arquivo de contratos aprovados para começar.")
    st.stop()  # interrompe a execução — nada abaixo roda sem o arquivo


# ─────────────────────────────────────────────────────────────────────────────
# LEITURA DO ARQUIVO PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────

try:
    # Lê a planilha pulando as 8 primeiras linhas (cabeçalho do sistema)
    df_aprovados = pd.read_excel(arquivo_aprovados, skiprows=8)

except Exception as erro:
    # Se der qualquer erro na leitura, mostra a mensagem e para
    st.error(f"❌ Erro ao ler o arquivo de contratos aprovados: {erro}")
    st.stop()

# Substitui células vazias por "-" para não aparecer "NaN" na tabela
df_aprovados = df_aprovados.fillna("-")

st.success("✅ Arquivo de contratos aprovados carregado!")


# ─────────────────────────────────────────────────────────────────────────────
# LEITURA DO ARQUIVO DO MÊS ANTERIOR (opcional)
# ─────────────────────────────────────────────────────────────────────────────

df_mes_anterior = None   # começa como None; só preenchemos se o arquivo for enviado

if arquivo_mes_anterior is not None:
    try:
        df_mes_anterior = pd.read_excel(arquivo_mes_anterior)
        df_mes_anterior = df_mes_anterior.fillna("-")
        st.success("✅ Arquivo do mês anterior carregado!")

    except Exception as erro:
        st.warning(f"⚠️ Erro ao ler o arquivo do mês anterior: {erro}")


# ─────────────────────────────────────────────────────────────────────────────
# PROCESSAMENTO
# Passa o DataFrame bruto pela função de processamento.
# ─────────────────────────────────────────────────────────────────────────────

try:
    df_processado = processar_contratos(df_aprovados)

except Exception as erro:
    st.error(f"❌ Erro durante o processamento dos dados: {erro}")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# DEBUG — Colunas disponíveis
# Recolhido por padrão; útil para diagnosticar problemas.
# ─────────────────────────────────────────────────────────────────────────────

with st.expander("🔍 Ver colunas encontradas no arquivo"):
    st.write(df_processado.columns.tolist())


# ─────────────────────────────────────────────────────────────────────────────
# EXIBIÇÃO — CONTRATOS APROVADOS
# Define quais colunas mostrar e exibe apenas as que existem no DataFrame.
# ─────────────────────────────────────────────────────────────────────────────

# Lista de colunas que queremos exibir, na ordem desejada
colunas_para_exibir = [
    'BOLETA',
    'OPERACAO',
    'FONTE',
    'PARTE',
    'CONTRAPARTE',
    'CP/LP',
    'SUBMERCADO',
    'MONTANTE MWh',
    'MONTANTE MWm',
    'CLIQ PARADIGMA',
    'MODULACAO WBC',
    'MOD MIN',
    'MOD MAX',
]

# Filtra só as colunas que realmente existem (evita erro de coluna não encontrada)
colunas_existentes = [
    col for col in colunas_para_exibir
    if col in df_processado.columns
]

# Avisa se alguma coluna esperada não existir
colunas_faltando = set(colunas_para_exibir) - set(colunas_existentes)
if colunas_faltando:
    st.warning(
        f"⚠️ Colunas não encontradas para exibição: "
        f"{', '.join(sorted(colunas_faltando))}"
    )

# Exibe a tabela de contratos aprovados
st.subheader("Contratos Aprovados")
st.dataframe(
    df_processado[colunas_existentes],
    hide_index=True,          # esconde o índice numérico das linhas
    use_container_width=True  # ocupa toda a largura disponível
)


# ─────────────────────────────────────────────────────────────────────────────
# EXIBIÇÃO — CONTRATOS DO MÊS ANTERIOR (quando enviado)
# TODO: implementar comparação entre df_processado e df_mes_anterior
# ─────────────────────────────────────────────────────────────────────────────

if df_mes_anterior is not None:
    st.subheader("Contratos Mês Anterior")
    st.dataframe(
        df_mes_anterior,
        hide_index=True,
        use_container_width=True
    )
