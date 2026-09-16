"""
Dashboard de Suporte N1 - Telecom (v2)
======================================
Ferramenta web de relatórios automáticos e dashboard interativo para
squads de atendimento Nível 1 (N1) de Telecom.

Principais melhorias desta versão:
  - Reconhece automaticamente as colunas reais do export (id/protocolo,
    Assunto, Data, Status, Analista, etc.) via mapeamento por sinônimos.
  - KPIs e gráficos se adaptam ao que existe na base.
  - Mantém compatibilidade retroativa: se um dia o export trouxer
    Tempo_Espera_Min, Tempo_Atendimento_Min e/ou Resolvido_N1, esses
    indicadores (TME, TMA, FCR) são calculados automaticamente.

Como executar:
    pip install streamlit pandas plotly openpyxl
    streamlit run dashboard_n1_telecom.py
"""

import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import streamlit as st


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO GERAL
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard N1 - Telecom",
    page_icon="📡",
    layout="wide",
)


# ---------------------------------------------------------------------------
# MAPEAMENTO DE CAMPOS CANÔNICOS -> SINÔNIMOS ACEITOS
# ---------------------------------------------------------------------------
# A chave é o "nome canônico" usado internamente pelo dashboard.
# Os valores são listas de padrões (regex, case-insensitive) que serão
# testados contra os nomes das colunas do arquivo carregado.
FIELD_SYNONYMS = {
    "Data": [
        r"^data(\s|_|$)", r"^data.*hora", r"^data.*abertura", r"^abertura",
        r"^criado", r"^created", r"^timestamp", r"^dt_", r"_dt$", r"^inicio",
        r"^início", r"^solicitacao", r"^solicitação",
    ],
    "Chamado": [
        r"^chamado", r"^ticket", r"^id", r"^protocolo", r"^numero", r"^número",
        r"^os$", r"^ordem", r"^codigo", r"^código", r"^interacao", r"^interação",
    ],
    "Motivo": [
        r"^motivo", r"^assunto", r"^categoria", r"^tipo", r"^topico", r"^tópico",
        r"^classificacao", r"^classificação", r"^servico", r"^serviço",
        r"^descricao", r"^descrição", r"^problema",
    ],
    "Status": [
        r"^status", r"^situacao", r"^situação", r"^estado", r"^state",
    ],
    "Analista": [
        r"^analista", r"^atendente", r"^operador", r"^agente", r"^responsavel",
        r"^responsável", r"^usuario", r"^usuário", r"^owner", r"^assigned",
        r"^tecnico", r"^técnico",
    ],
    "Resolvido_N1": [
        r"^resolvido", r"^resolvido_n1", r"^fcr", r"^resolucao", r"^resolução",
        r"^solucionado", r"^primeiro_nivel", r"^primeiro_nível",
    ],
    "Tempo_Espera_Min": [
        r"^tempo_espera", r"^tempo_espera_min", r"^tme", r"^espera",
        r"^wait_time", r"^waiting", r"^fila",
    ],
    "Tempo_Atendimento_Min": [
        r"^tempo_atendimento", r"^tempo_atendimento_min", r"^tma",
        r"^duracao", r"^duração", r"^handle_time", r"^aht",
    ],
    "Prioridade": [
        r"^prioridade", r"^priority", r"^severidade", r"^urgencia", r"^urgência",
    ],
    "Canal": [
        r"^canal", r"^channel", r"^origem", r"^fonte", r"^midia", r"^mídia",
    ],
    "Cliente": [
        r"^cliente", r"^customer", r"^contrato", r"^assinante", r"^conta",
    ],
}

# Campos mínimos obrigatórios para o dashboard funcionar.
# (Data é indispensável; os demais são opcionais e habilitam KPIs/gráficos.)
CAMPO_MINIMO_OBRIGATORIO = "Data"


# ---------------------------------------------------------------------------
# FUNÇÕES DE APOIO À DETECÇÃO DE COLUNAS
# ---------------------------------------------------------------------------
def _normalizar(texto: str) -> str:
    """Remove acentos, coloca em minúsculas e troca separadores por '_'."""
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.strip().lower()
    texto = re.sub(r"[\s\-\.]+", "_", texto)
    return texto


def detectar_colunas(df: pd.DataFrame) -> dict:
    """
    Retorna um dicionário {campo_canonico: nome_real_da_coluna}.
    Caso o campo não exista no arquivo, ele é omitido do dicionário.
    """
    mapa = {}
    colunas_norm = {col: _normalizar(col) for col in df.columns}

    for campo, padroes in FIELD_SYNONYMS.items():
        for col_original, col_norm in colunas_norm.items():
            if col_original in mapa.values():
                continue
            for padrao in padroes:
                if re.search(padrao, col_norm):
                    mapa[campo] = col_original
                    break
            if campo in mapa:
                break
    return mapa


# ---------------------------------------------------------------------------
# LEITURA E PREPARAÇÃO DOS DADOS
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Processando arquivo carregado...")
def load_data(file) -> pd.DataFrame:
    nome = file.name.lower()
    if nome.endswith(".csv"):
        try:
            df = pd.read_csv(file, sep=None, engine="python")
        except Exception:
            file.seek(0)
            df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    return df


def preparar_dados(df_raw: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    """
    Renomeia as colunas detectadas para nomes canônicos, converte tipos e
    limpa a base. As colunas originais que NÃO foram mapeadas são mantidas
    (úteis para auditoria).
    """
    df = df_raw.copy()

    # 1) Renomeia colunas detectadas para nomes canônicos
    df = df.rename(columns={v: k for k, v in mapa.items()})

    # 2) Converte Data
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
        invalidas = int(df["Data"].isna().sum())
        if invalidas > 0:
            st.warning(
                f"⚠️ {invalidas} registro(s) com data/hora inválida foram "
                "descartados automaticamente da análise."
            )
            df = df.dropna(subset=["Data"])

    # 3) Converte numéricos opcionais
    for col in ("Tempo_Espera_Min", "Tempo_Atendimento_Min"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # 4) Padroniza textos
    for col in ("Motivo", "Status", "Analista", "Prioridade", "Canal"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "Resolvido_N1" in df.columns:
        df["Resolvido_N1"] = (
            df["Resolvido_N1"].astype(str).str.strip().str.capitalize()
        )

    df = df.sort_values("Data").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# GRÁFICO DE EVOLUÇÃO TEMPORAL (ADAPTATIVO)
# ---------------------------------------------------------------------------
def build_time_evolution_chart(df_periodo, period_option, data_inicio, data_fim):
    df_local = df_periodo.copy()
    delta_dias = (data_fim.normalize() - data_inicio.normalize()).days

    def _linha(contagem, titulo_x, formato_x=None):
        fig = px.line(contagem, x="Periodo", y="Chamados", markers=True)
        fig.update_xaxes(title_text=titulo_x, tickformat=formato_x)
        fig.update_traces(line=dict(width=3), fill="tozeroy",
                          fillcolor="rgba(0,123,255,0.12)")
        fig.update_layout(
            yaxis_title="Nº de Chamados",
            hovermode="x unified",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        return fig

    def agrega_hora(inicio, fim):
        df_local["Agrupador"] = df_local["Data"].dt.floor("H")
        faixa = pd.date_range(start=inicio.floor("H"), end=fim.floor("H"), freq="H")
        contagem = (df_local.groupby("Agrupador").size()
                    .reindex(faixa, fill_value=0).reset_index())
        contagem.columns = ["Periodo", "Chamados"]
        return _linha(contagem, "Hora do Dia", "%H:%M")

    def agrega_dia(inicio, fim):
        df_local["Agrupador"] = df_local["Data"].dt.normalize()
        faixa = pd.date_range(start=inicio.normalize(), end=fim.normalize(), freq="D")
        contagem = (df_local.groupby("Agrupador").size()
                    .reindex(faixa, fill_value=0).reset_index())
        contagem.columns = ["Periodo", "Chamados"]
        return _linha(contagem, "Dia", "%d/%m")

    def agrega_semana():
        df_local["Agrupador"] = (df_local["Data"].dt.to_period("W")
                                 .apply(lambda r: r.start_time))
        contagem = df_local.groupby("Agrupador").size().reset_index()
        contagem.columns = ["Periodo", "Chamados"]
        contagem = contagem.sort_values("Periodo")
        contagem["Rotulo"] = contagem["Periodo"].apply(
            lambda d: f"Semana de {d.strftime('%d/%m')}"
        )
        fig = px.line(contagem, x="Rotulo", y="Chamados", markers=True)
        fig.update_xaxes(title_text="Semana")
        fig.update_traces(line=dict(width=3), fill="tozeroy",
                          fillcolor="rgba(0,123,255,0.12)")
        fig.update_layout(
            yaxis_title="Nº de Chamados",
            hovermode="x unified",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        return fig

    if period_option == "Diário (Hoje)":
        return agrega_hora(data_inicio, data_fim)
    if period_option == "Semanal (Últimos 7 dias)":
        return agrega_dia(data_inicio, data_fim)
    if period_option == "Mensal (Últimos 30 dias)":
        return agrega_semana()

    # Personalizado
    if delta_dias <= 1:
        return agrega_hora(data_inicio, data_fim)
    if delta_dias <= 31:
        return agrega_dia(data_inicio, data_fim)
    return agrega_semana()


# ---------------------------------------------------------------------------
# CABEÇALHO
# ---------------------------------------------------------------------------
st.title("📡 Dashboard de Suporte N1 - Telecom")
st.caption(
    "Painel automático de indicadores, evolução temporal e auditoria para "
    "squads de atendimento Nível 1."
)


# ---------------------------------------------------------------------------
# UPLOAD
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Configurações do Relatório")
uploaded_file = st.sidebar.file_uploader(
    "📂 Carregue a base de atendimentos (.csv ou .xlsx)",
    type=["csv", "xlsx"],
)

if uploaded_file is None:
    st.info(
        "👋 **Bem-vindo(a)!** Para começar, carregue o arquivo de atendimentos "
        "(.csv ou .xlsx) na barra lateral à esquerda."
    )
    st.markdown("#### 🔎 O dashboard reconhece automaticamente colunas como:")
    st.dataframe(
        pd.DataFrame(
            {
                "Campo interno": list(FIELD_SYNONYMS.keys()),
                "Exemplos de nomes aceitos": [
                    "Data, Data_Abertura, Criado_Em, Timestamp",
                    "Chamado, Ticket, ID, Protocolo, Número, OS",
                    "Motivo, Assunto, Categoria, Tipo, Tópico, Problema",
                    "Status, Situação, Estado",
                    "Analista, Atendente, Operador, Agente, Técnico, Responsável",
                    "Resolvido_N1, FCR, Solucionado, Primeiro_Nivel (opcional)",
                    "Tempo_Espera_Min, TME, Espera, Wait_Time (opcional)",
                    "Tempo_Atendimento_Min, TMA, Duração, AHT (opcional)",
                    "Prioridade, Severidade, Urgência (opcional)",
                    "Canal, Channel, Origem, Mídia (opcional)",
                ],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.caption(
        "Se um campo opcional não existir, o KPI/gráfico correspondente "
        "simplesmente não é exibido — o restante continua funcionando."
    )
    st.stop()


# ---------------------------------------------------------------------------
# LEITURA + DETECÇÃO DE COLUNAS
# ---------------------------------------------------------------------------
try:
    df_raw = load_data(uploaded_file)
except Exception as erro:
    st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes técnicos: {erro}")
    st.stop()

if df_raw.empty:
    st.error("❌ O arquivo carregado está vazio.")
    st.stop()

mapa_colunas = detectar_colunas(df_raw)

if CAMPO_MINIMO_OBRIGATORIO not in mapa_colunas:
    st.error(
        "❌ Não foi possível identificar uma coluna de **Data** no arquivo. "
        "Verifique se existe alguma coluna com nome semelhante a "
        "`Data`, `Data_Abertura`, `Criado_Em` ou `Timestamp`."
    )
    st.stop()

with st.sidebar.expander("🧭 Colunas reconhecidas"):
    for campo, col_real in mapa_colunas.items():
        st.markdown(f"- **{campo}** ← `{col_real}`")
    nao_mapeadas = [c for c in df_raw.columns if c not in mapa_colunas.values()]
    if nao_mapeadas:
        st.markdown("**Não utilizadas:**")
        st.markdown(", ".join(f"`{c}`" for c in nao_mapeadas))

df = preparar_dados(df_raw, mapa_colunas)

if df.empty:
    st.error("❌ Após a validação, nenhum registro válido restou para análise.")
    st.stop()


# ---------------------------------------------------------------------------
# SELETOR DE PERÍODO
# ---------------------------------------------------------------------------
st.sidebar.markdown("---")
period_option = st.sidebar.selectbox(
    "📅 Período do Relatório",
    [
        "Diário (Hoje)",
        "Semanal (Últimos 7 dias)",
        "Mensal (Últimos 30 dias)",
        "Período Personalizado",
    ],
)

hoje = pd.Timestamp(datetime.now().date())

if period_option == "Diário (Hoje)":
    data_inicio = hoje
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)
elif period_option == "Semanal (Últimos 7 dias)":
    data_inicio = hoje - timedelta(days=6)
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)
elif period_option == "Mensal (Últimos 30 dias)":
    data_inicio = hoje - timedelta(days=29)
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)
else:
    st.sidebar.markdown("**Selecione o intervalo desejado:**")
    col_a, col_b = st.sidebar.columns(2)
    with col_a:
        ini_input = st.date_input("Data Início", value=(hoje - timedelta(days=7)).date())
    with col_b:
        fim_input = st.date_input("Data Fim", value=hoje.date())
    if ini_input > fim_input:
        st.sidebar.error("⚠️ A data de início não pode ser posterior à data de fim.")
        st.stop()
    data_inicio = pd.Timestamp(ini_input)
    data_fim = pd.Timestamp(fim_input) + timedelta(days=1) - timedelta(seconds=1)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"🗓️ Exibindo de **{data_inicio.strftime('%d/%m/%Y %H:%M')}** "
    f"até **{data_fim.strftime('%d/%m/%Y %H:%M')}**"
)


# ---------------------------------------------------------------------------
# FILTRAGEM
# ---------------------------------------------------------------------------
df_filtrado = df[(df["Data"] >= data_inicio) & (df["Data"] <= data_fim)].copy()

if df_filtrado.empty:
    st.warning(
        "⚠️ Nenhum chamado encontrado no período selecionado. Ajuste o filtro "
        "na barra lateral ou verifique se a base contém registros nesta janela."
    )
    st.stop()


# ---------------------------------------------------------------------------
# KPIs ADAPTATIVOS
# ---------------------------------------------------------------------------
st.markdown("### 📊 Indicadores do Período")

total_chamados = len(df_filtrado)
kpis = []

kpis.append(("📞 Total de Chamados",
             f"{total_chamados:,}".replace(",", ".")))

# FCR (só se Resolvido_N1 existir)
if "Resolvido_N1" in df_filtrado.columns:
    resolvidos = int((df_filtrado["Resolvido_N1"] == "Sim").sum())
    taxa_fcr = (resolvidos / total_chamados * 100) if total_chamados else 0.0
    kpis.append(("✅ Taxa de FCR (N1)", f"{taxa_fcr:.1f}%"))

# TME
if "Tempo_Espera_Min" in df_filtrado.columns:
    tme = df_filtrado["Tempo_Espera_Min"].mean()
    kpis.append(("⏱️ TME Médio", f"{tme:.1f} min" if pd.notna(tme) else "N/A"))

# TMA
if "Tempo_Atendimento_Min" in df_filtrado.columns:
    tma = df_filtrado["Tempo_Atendimento_Min"].mean()
    kpis.append(("🎧 TMA Médio", f"{tma:.1f} min" if pd.notna(tma) else "N/A"))

# Analistas ativos (só se Analista existir)
if "Analista" in df_filtrado.columns:
    n_analistas = df_filtrado["Analista"].nunique()
    kpis.append(("👥 Analistas Ativos", f"{n_analistas}"))

# Status "Aberto"/"Pendente" (heurística simples)
if "Status" in df_filtrado.columns:
    status_series = df_filtrado["Status"].str.lower()
    em_aberto = status_series.str.contains(
        r"aberto|pendente|em andamento|aguardando|andamento",
        regex=True, na=False
    ).sum()
    kpis.append(("📬 Chamados em Aberto", f"{int(em_aberto)}"))

# Renderiza KPIs em grade responsiva
COLS_POR_LINHA = 4
for i in range(0, len(kpis), COLS_POR_LINHA):
    linha = kpis[i:i + COLS_POR_LINHA]
    cols = st.columns(len(linha))
    for col, (titulo, valor) in zip(cols, linha):
        col.metric(titulo, valor)

st.markdown("---")


# ---------------------------------------------------------------------------
# GRÁFICOS ADAPTATIVOS
# ---------------------------------------------------------------------------
grafico_motivo_disponivel = "Motivo" in df_filtrado.columns
grafico_status_disponivel = "Status" in df_filtrado.columns
grafico_analista_disponivel = "Analista" in df_filtrado.columns

col_graf1, col_graf2 = st.columns(2)

with col_graf1:
    if grafico_motivo_disponivel:
        st.markdown("#### 🔎 Principais Motivos de Contato")
        motivo_counts = df_filtrado["Motivo"].value_counts().head(15).reset_index()
        motivo_counts.columns = ["Motivo", "Quantidade"]
        fig_motivos = px.bar(
            motivo_counts, x="Quantidade", y="Motivo", orientation="h",
            text="Quantidade", color="Quantidade",
            color_continuous_scale="Blues",
        )
        fig_motivos.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Nº de Chamados", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        fig_motivos.update_traces(textposition="outside")
        st.plotly_chart(fig_motivos, use_container_width=True)
    elif grafico_status_disponivel:
        st.markdown("#### 📌 Distribuição por Status")
        st_counts = df_filtrado["Status"].value_counts().reset_index()
        st_counts.columns = ["Status", "Quantidade"]
        fig_st = px.pie(st_counts, names="Status", values="Quantidade", hole=0.5)
        fig_st.update_layout(margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_st, use_container_width=True)
    else:
        st.info("ℹ️ Nenhuma coluna de Motivo/Assunto/Categoria foi reconhecida.")

with col_graf2:
    rotulo_periodo = period_option.split(" ")[0]
    st.markdown(f"#### 📈 Evolução de Volume — Visão {rotulo_periodo}")
    fig_evolucao = build_time_evolution_chart(
        df_filtrado, period_option, data_inicio, data_fim
    )
    st.plotly_chart(fig_evolucao, use_container_width=True)


# Linha extra de gráficos: Status + Analista (se disponíveis)
col_graf3, col_graf4 = st.columns(2)

with col_graf3:
    if grafico_status_disponivel and grafico_motivo_disponivel:
        st.markdown("#### 📌 Distribuição por Status")
        st_counts = df_filtrado["Status"].value_counts().reset_index()
        st_counts.columns = ["Status", "Quantidade"]
        fig_st = px.pie(st_counts, names="Status", values="Quantidade", hole=0.5)
        fig_st.update_layout(margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig_st, use_container_width=True)

with col_graf4:
    if grafico_analista_disponivel:
        st.markdown("#### 🧑‍💻 Top Analistas por Volume")
        an_counts = df_filtrado["Analista"].value_counts().head(15).reset_index()
        an_counts.columns = ["Analista", "Quantidade"]
        fig_an = px.bar(
            an_counts, x="Quantidade", y="Analista", orientation="h",
            text="Quantidade", color="Quantidade",
            color_continuous_scale="Greens",
        )
        fig_an.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Nº de Chamados", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        fig_an.update_traces(textposition="outside")
        st.plotly_chart(fig_an, use_container_width=True)

st.markdown("---")


# ---------------------------------------------------------------------------
# TABELA DE AUDITORIA
# ---------------------------------------------------------------------------
st.markdown("### 🗂️ Auditoria de Chamados (Dados Brutos)")
st.caption(
    "Utilize a tabela abaixo para conferência detalhada dos registros do "
    "período filtrado. Os dados podem ser exportados em CSV."
)

df_exibicao = df_filtrado.sort_values("Data", ascending=False).reset_index(drop=True)
st.dataframe(df_exibicao, use_container_width=True, height=420)

csv_export = df_exibicao.to_csv(index=False, sep=";").encode("utf-8-sig")
st.download_button(
    label="⬇️ Exportar CSV do Período Filtrado",
    data=csv_export,
    file_name=f"auditoria_n1_{data_inicio.date()}_a_{data_fim.date()}.csv",
    mime="text/csv",
)
