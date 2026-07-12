"""
Relatório de Manutenção por Setor - Fábrica de Açúcar
=======================================================
App em Streamlit que:
  1. Lê a planilha de dados de compras/serviços.
  2. Unifica os nomes da coluna 'LOCAL SERV' (setores duplicados/parecidos).
  3. Remove o setor 'Outros'.
  4. Exibe, para cada setor:
       - um slider de percentual de manutenção (0% a 10%) próprio, logo
         acima do expander do setor, exibindo o símbolo de "%";
       - um expander (+) com a lista de equipamentos (Desc. Prod. / Valor
         Total), onde é possível marcar 'Remover' para desconsiderar um
         equipamento do cálculo. A linha NÃO desaparece: ela continua
         visível, mas o 'Valor Total' daquele item passa a ser R$ 0,00.
       - a lista de itens removidos é guardada em st.session_state (fonte
         de verdade única), então marcar um novo item NUNCA "esquece" os
         itens já removidos anteriormente — nem no mesmo setor, nem em
         outros setores. Um item só volta ao normal se o usuário
         desmarcar manualmente a caixa 'Remover?'.
       - o valor total do setor e a projeção de manutenção recalculados
         instantaneamente, sem alterar a planilha original.

Para rodar:
    streamlit run app.py
"""

import unicodedata

import pandas as pd
import streamlit as st

# ----------------------------------------------------------------------
# Configuração da página
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Relatório de Manutenção por Setor",
    page_icon="🏭",
    layout="wide",
)

CAMINHO_PADRAO = "DadosFabAcucar.xlsx"

# ----------------------------------------------------------------------
# 1. Unificação dos nomes de 'LOCAL SERV'
# ----------------------------------------------------------------------
# A planilha traz o mesmo setor grafado de formas diferentes (com/sem
# acento, maiúsculas/minúsculas, espaços extras). Para agrupar de forma
# robusta, normalizamos a string (removendo acentos, espaços e deixando
# tudo em maiúsculas) e comparamos com uma chave "canônica".


def normalizar(texto: str) -> str:
    """Remove acentos, espaços extras e deixa em maiúsculas para comparação."""
    if texto is None:
        return ""
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    # colapsa múltiplos espaços em um só
    texto = " ".join(texto.split())
    return texto


# Mapa: chave normalizada (sem acento) -> nome final que será exibido no app
MAPA_UNIFICACAO = {
    # Centrífuga
    "CENTRIFUGA": "Centrífuga",
    "CENTRIFUGACAO": "Centrífuga",
    # Cozedores
    "COZEDORES": "Cozedores",
    "COZIMENTO": "Cozedores",
    # Fábrica de Açúcar
    "FABRICA DE ACUCAR": "Fábrica de Açúcar",
    "FABRICACAO DE ACUCAR": "Fábrica de Açúcar",
    # Geração de Energia
    "GERACAO DE ENERGIA": "Geração de Energia",
    "GERACAO DE ENERGIA ELETRICA- CASA DE FORCA 2": "Geração de Energia",
    "GERACAO DE ENERGIA ELETRICA-2 (AMPLIACAO)": "Geração de Energia",
    "GERACAO DE ENERGIA ELETRICA CASA DE FORCA": "Geração de Energia",
    "GERACAO DE ENERGIA 2": "Geração de Energia",
    # Setores que permanecem como estão
    "CRISTALIZADORES": "Cristalizadores",
    "LABORATORIO DE CONTROLE INDUSTRIAL": "Laboratório de Controle Industrial",
    "REFINARIA": "Refinaria",
    "SECADOR": "Secador",
    "TRATAMENTO DE CALDO": "Tratamento de Caldo",
    # Setor a ser removido do relatório
    "OUTROS": None,
}


def unificar_local(texto: str) -> str:
    """Aplica o mapa de unificação; se não encontrar, mantém o texto original
    (apenas limpo de espaços extras), garantindo que setores não previstos
    na regra não sejam perdidos silenciosamente."""
    chave = normalizar(texto)
    if chave in MAPA_UNIFICACAO:
        return MAPA_UNIFICACAO[chave]
    # setor não mapeado explicitamente -> mantém como está (só limpa espaços)
    return " ".join(str(texto).strip().split()).title() if texto else texto


# ----------------------------------------------------------------------
# 2. Carregamento e preparação dos dados
# ----------------------------------------------------------------------
@st.cache_data(show_spinner="Lendo planilha...")
def carregar_dados(arquivo) -> pd.DataFrame:
    df = pd.read_excel(arquivo)

    colunas_necessarias = ["LOCAL SERV", "Valor Total", "Desc. Prod."]
    faltando = [c for c in colunas_necessarias if c not in df.columns]
    if faltando:
        raise ValueError(
            f"As colunas obrigatórias não foram encontradas na planilha: {faltando}"
        )

    df = df[colunas_necessarias].copy()

    # Aplica a unificação dos setores
    df["LOCAL SERV"] = df["LOCAL SERV"].apply(unificar_local)

    # Remove o setor 'Outros' (unificar_local devolve None para ele)
    df = df[df["LOCAL SERV"].notna()]

    # Garante tipos corretos
    df["Valor Total"] = pd.to_numeric(df["Valor Total"], errors="coerce").fillna(0)
    df["Desc. Prod."] = df["Desc. Prod."].astype(str).str.strip()

    # ID único e estável para cada linha (usado para controlar remoções)
    df = df.reset_index(drop=True)
    df["ID"] = df.index

    return df


# ----------------------------------------------------------------------
# 3. Interface
# ----------------------------------------------------------------------
st.title("🏭 Relatório de Manutenção por Setor")
st.caption(
    "Simulação de projeção de manutenção por setor, a partir dos equipamentos "
    "cadastrados. Remoções e percentuais aqui não alteram a planilha original."
)

arquivo_upado = st.file_uploader(
    "Envie a planilha (.xlsx). Se não enviar, o app tentará usar "
    f"'{CAMINHO_PADRAO}' na mesma pasta do script.",
    type=["xlsx"],
)

fonte_dados = arquivo_upado if arquivo_upado is not None else CAMINHO_PADRAO

try:
    df = carregar_dados(fonte_dados)
except FileNotFoundError:
    st.warning(
        f"Nenhum arquivo enviado e '{CAMINHO_PADRAO}' não foi encontrado na "
        "pasta do script. Envie a planilha acima para continuar."
    )
    st.stop()
except ValueError as e:
    st.error(str(e))
    st.stop()

setores = sorted(df["LOCAL SERV"].unique())

# ----------------------------------------------------------------------
# Fonte única de verdade para os itens removidos: um conjunto de IDs
# guardado em st.session_state, que persiste durante toda a sessão do
# usuário (independente de resets internos do st.data_editor).
# ----------------------------------------------------------------------
if "itens_removidos" not in st.session_state:
    st.session_state["itens_removidos"] = set()

st.divider()

valor_total_geral = 0.0
projecao_total_geral = 0.0

# Uma seção por setor: slider (na página principal) + expander com a tabela
for setor in setores:
    df_setor = df[df["LOCAL SERV"] == setor][
        ["ID", "Desc. Prod.", "Valor Total"]
    ].reset_index(drop=True)

    editor_key = f"editor_{setor}"

    # Monta a tabela a partir da fonte única de verdade (session_state):
    # marca 'Remover' e zera o Valor Total para qualquer ID que já esteja
    # no conjunto de itens removidos — não importa se foi marcado agora,
    # neste setor, ou em qualquer interação anterior.
    df_base = df_setor.copy()
    df_base["Remover"] = df_base["ID"].isin(st.session_state["itens_removidos"])
    df_base.loc[df_base["Remover"], "Valor Total"] = 0.0

    # -----------------------------------------------------------------
    # Slider de percentual — na página principal, ACIMA do expander,
    # independente para cada setor. format="%d%%" exibe o símbolo de %.
    # -----------------------------------------------------------------
    st.markdown(f"**{setor}**")
    percentual = st.slider(
        f"Percentual de manutenção — {setor}",
        min_value=0,
        max_value=10,
        value=5,
        step=1,
        format="%d%%",
        key=f"pct_{setor}",
        label_visibility="collapsed",
    )
    st.caption(f"Percentual de manutenção aplicado: **{percentual}%**")

    with st.expander(f"📦 {setor}  —  {len(df_base)} equipamento(s)"):

        st.markdown(
            "Marque **Remover** para desconsiderar o equipamento do cálculo. "
            "O item continua listado, mas o Valor Total dele passa a ser R$ 0,00. "
            "Desmarque a caixa para trazê-lo de volta ao cálculo."
        )

        editado = st.data_editor(
            df_base,
            key=editor_key,
            hide_index=True,
            use_container_width=True,
            disabled=["ID", "Desc. Prod.", "Valor Total"],
            column_config={
                "ID": None,  # oculta a coluna de controle interno
                "Desc. Prod.": st.column_config.TextColumn("Equipamento (Desc. Prod.)"),
                "Valor Total": st.column_config.NumberColumn(
                    "Valor Total (R$)", format="R$ %.2f"
                ),
                "Remover": st.column_config.CheckboxColumn("Remover?"),
            },
        )

        # Sincroniza a interação desta rodada de volta para a fonte única
        # de verdade: marcou -> adiciona o ID ao conjunto; desmarcou ->
        # remove o ID do conjunto. Assim, o estado nunca é perdido ao
        # marcar outros itens (no mesmo setor ou em outros setores), e só
        # "reseta" quando o usuário desmarca manualmente a caixa.
        marcados_agora = set(editado.loc[editado["Remover"], "ID"])
        desmarcados_agora = set(editado.loc[~editado["Remover"], "ID"])
        st.session_state["itens_removidos"] |= marcados_agora
        st.session_state["itens_removidos"] -= desmarcados_agora

        # Recalcula o Valor Total já refletindo a sincronização acima.
        valor_total_setor = editado["Valor Total"].sum()
        qtd_removidos = int(editado["Remover"].sum())
        projecao_setor = valor_total_setor * (percentual / 100)

        valor_total_geral += valor_total_setor
        projecao_total_geral += projecao_setor

        m1, m2, m3 = st.columns(3)
        m1.metric("Equip. removidos (zerados)", qtd_removidos)
        m2.metric("Valor total do setor", f"R$ {valor_total_setor:,.2f}")
        m3.metric(
            f"Projeção de manutenção ({percentual}%)",
            f"R$ {projecao_setor:,.2f}",
        )

    st.divider()

# ----------------------------------------------------------------------
# 4. Totais gerais (considerando remoções e percentuais de todos os setores)
# ----------------------------------------------------------------------
st.subheader("📊 Totais gerais")
c1, c2 = st.columns(2)
c1.metric("Valor total geral (todos os setores)", f"R$ {valor_total_geral:,.2f}")
c2.metric("Projeção de manutenção total", f"R$ {projecao_total_geral:,.2f}")
