"""
Dashboard de Suporte N1 - Telecom (v3)
======================================
Adaptado ao export real da plataforma (id, protocolo, Solicitante,
Assunto, Data, Setor, Analista, Urgência, Interno, Status, Data Encerrado).

Recursos:
  - Reconhecimento automático de colunas por sinônimos.
  - KPIs adaptativos: só aparece o que faz sentido com os dados.
  - TMA real calculado quando há "Data Encerrado".
  - Taxa de Resolução N1, Taxa de Cancelamento, Ranking de Analistas,
    Top Assuntos, Distribuição de Status, Evolução temporal e
    Heatmap Hora × Dia da Semana.
  - Compatibilidade retroativa com exports que tragam Tempo_Espera_Min,
    Tempo_Atendimento_Min e Resolvido_N1.

Como executar:
    pip install streamlit pandas plotly openpyxl
    streamlit run dashboard_n1_telecom.py
"""

import re
import unicodedata
from datetime import datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
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
# MAPEAMENTO DE CAMPOS CANÔNICOS -> SINÔNIMOS
# ---------------------------------------------------------------------------
FIELD_SYNONYMS = {
    "Data": [
        r"^data$", r"^data_abertura", r"^data_criacao", r"^abertura",
        r"^criado_em", r"^timestamp", r"^dt_", r"_dt$", r"^inicio",
    ],
    "Data_Encerrado": [
        r"^data_encerrado", r"^data_fechamento", r"^data_conclusao",
        r"^data_finalizacao", r"^fechado_em", r"^closed_at", r"^encerrado",
    ],
    "Chamado": [
        r"^id$", r"^chamado", r"^ticket", r"^numero", r"^os$", r"^codigo",
    ],
    "Protocolo": [
        r"^protocolo", r"^protocol$", r"^numero_protocolo",
    ],
    "Solicitante": [
        r"^solicitante", r"^cliente", r"^customer", r"^assinante", r"^conta",
    ],
    "Motivo": [
        r"^assunto", r"^motivo", r"^categoria", r"^tipo", r"^topico",
        r"^classificacao", r"^problema",
    ],
    "Setor": [
        r"^setor", r"^departamento", r"^fila", r"^grupo",
    ],
    "Analista": [
        r"^analista", r"^atendente", r"^operador", r"^agente",
        r"^responsavel", r"^tecnico", r"^owner", r"^assigned",
    ],
    "Status": [
        r"^status$", r"^situacao", r"^estado", r"^state",
    ],
    "Urgencia": [
        r"^urgencia", r"^prioridade", r"^priority", r"^severidade",
    ],
    "Interno": [
        r"^interno", r"^internal", r"^is_internal",
    ],
    "Resolvido_N1": [
        r"^resolvido", r"^resolvido_n1", r"^fcr", r"^solucionado",
    ],
    "Tempo_Espera_Min": [
        r"^tempo_espera", r"^tme", r"^espera", r"^wait_time", r"^fila",
    ],
    "Tempo_Atendimento_Min": [
        r"^tempo_atendimento", r"^tma", r"^duracao", r"^aht", r"^handle_time",
    ],
}

CAMPO_MINIMO_OBRIGATORIO = "Data"


# ---------------------------------------------------------------------------
# NORMALIZAÇÃO E DETECÇÃO DE COLUNAS
# ---------------------------------------------------------------------------
def _normalizar(texto) -> str:
    if texto is None:
        return ""
    texto = unicodedata.normalize("NFKD", str(texto))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.strip().lower()
    texto = re.sub(r"[\s\-\.]+", "_", texto)
    return texto


def detectar_colunas(df: pd.DataFrame) -> dict:
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
# LEITURA E PREPARAÇÃO
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Processando arquivo carregado...")
def load_data(file) -> pd.DataFrame:
    nome = file.name.lower()
    if nome.endswith(".csv"):
        # detecta separador automaticamente (vírgula ou ponto e vírgula)
        try:
            df = pd.read_csv(file, sep=None, engine="python")
        except Exception:
            file.seek(0)
            df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    return df


def preparar_dados(df_raw: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    df = df_raw.copy()
    df = df.rename(columns={v: k for k, v in mapa.items()})

    # --- Datas -------------------------------------------------------------
    for col_data in ("Data", "Data_Encerrado"):
        if col_data in df.columns:
            df[col_data] = pd.to_datetime(
                df[col_data], errors="coerce", dayfirst=True
            )

    if "Data" in df.columns:
        invalidas = int(df["Data"].isna().sum())
        if invalidas > 0:
            st.warning(
                f"⚠️ {invalidas} registro(s) com data inválida foram descartados."
            )
            df = df.dropna(subset=["Data"])

    # --- Tempos ------------------------------------------------------------
    for col in ("Tempo_Espera_Min", "Tempo_Atendimento_Min"):
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # --- TMA derivado de Data_Encerrado - Data -----------------------------
    if "Data_Encerrado" in df.columns and "Data" in df.columns:
        delta = (df["Data_Encerrado"] - df["Data"]).dt.total_seconds() / 60.0
        # Só considera positivo (ignora negativos ou zero para "Em Atendimento")
        df["TMA_Calculado_Min"] = delta.where(delta > 0)

    # --- Textos ------------------------------------------------------------
    for col in ("Motivo", "Status", "Analista", "Urgencia", "Setor"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Motivo: capitaliza para agrupar "desbloqueio por confiança" e "DESBLOQUEIO..."
    if "Motivo" in df.columns:
        df["Motivo_Original"] = df["Motivo"]
        df["Motivo"] = (
            df["Motivo"]
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .str.title()
        )

    if "Resolvido_N1" in df.columns:
        df["Resolvido_N1"] = (
            df["Resolvido_N1"].astype(str).str.strip().str.capitalize()
        )

    df = df.sort_values("Data").reset_index(drop=True)
    return df


# ---------------------------------------------------------------------------
# GRÁFICO DE EVOLUÇÃO TEMPORAL
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

    if delta_dias <= 1:
        return agrega_hora(data_inicio, data_fim)
    if delta_dias <= 31:
        return agrega_dia(data_inicio, data_fim)
    return agrega_semana()


# ---------------------------------------------------------------------------
# HEATMAP HORA x DIA DA SEMANA
# ---------------------------------------------------------------------------
def build_heatmap_hora_dia(df_periodo):
    dias_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    df_h = df_periodo.copy()
    df_h["DiaSemana"] = df_h["Data"].dt.weekday
    df_h["Hora"] = df_h["Data"].dt.hour
    pivot = (df_h.groupby(["DiaSemana", "Hora"]).size()
             .unstack(fill_value=0)
             .reindex(index=range(7), columns=range(24), fill_value=0))
    pivot.index = [dias_pt[i] for i in pivot.index]

    fig = go.Figure(data=go.Heatmap(
        z=pivot.values,
        x=[f"{h:02d}h" for h in pivot.columns],
        y=pivot.index,
        colorscale="Blues",
        text=pivot.values,
        texttemplate="%{text}",
        hovertemplate="Dia: %{y}<br>Hora: %{x}<br>Chamados: %{z}<extra></extra>",
    ))
    fig.update_layout(
        xaxis_title="Hora",
        yaxis_title="",
        margin=dict(l=10, r=10, t=30, b=10),
        height=320,
    )
    return fig


# ---------------------------------------------------------------------------
# CABEÇALHO
# ---------------------------------------------------------------------------
st.title("📡 Dashboard de Suporte N1 - Telecom")
st.caption(
    "Painel automático de indicadores, evolução temporal e auditoria "
    "para squads de atendimento Nível 1."
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
        "👋 **Bem-vindo(a)!** Carregue o export de atendimentos N1 "
        "(.csv ou .xlsx) na barra lateral para começar."
    )
    st.markdown("#### 🔎 Colunas reconhecidas automaticamente:")
    st.dataframe(
        pd.DataFrame({
            "Campo interno": list(FIELD_SYNONYMS.keys()),
            "Exemplos de nomes aceitos": [
                "Data, Data_Abertura, Criado_Em, Abertura",
                "Data_Encerrado, Data_Fechamento, Closed_At",
                "ID, Chamado, Ticket, Número, OS, Código",
                "Protocolo, Numero_Protocolo",
                "Solicitante, Cliente, Assinante, Conta",
                "Assunto, Motivo, Categoria, Tipo, Tópico",
                "Setor, Departamento, Fila, Grupo",
                "Analista, Atendente, Operador, Agente, Técnico",
                "Status, Situação, Estado",
                "Urgência, Prioridade, Severidade",
                "Interno, Internal, Is_Internal",
                "Resolvido_N1, FCR, Solucionado (opcional)",
                "Tempo_Espera_Min, TME (opcional)",
                "Tempo_Atendimento_Min, TMA, Duração (opcional)",
            ],
        }),
        use_container_width=True, hide_index=True,
    )
    st.stop()


# ---------------------------------------------------------------------------
# LEITURA + DETECÇÃO
# ---------------------------------------------------------------------------
try:
    df_raw = load_data(uploaded_file)
except Exception as erro:
    st.error(f"❌ Não foi possível ler o arquivo. Detalhes: {erro}")
    st.stop()

if df_raw.empty:
    st.error("❌ Arquivo vazio.")
    st.stop()

mapa_colunas = detectar_colunas(df_raw)

if CAMPO_MINIMO_OBRIGATORIO not in mapa_colunas:
    st.error(
        "❌ Não foi possível identificar a coluna de **Data**. "
        "Verifique se existe algo como `Data`, `Data_Abertura` ou `Criado_Em`."
    )
    st.stop()

with st.sidebar.expander("🧭 Colunas reconhecidas", expanded=False):
    for campo, col_real in mapa_colunas.items():
        st.markdown(f"- **{campo}** ← `{col_real}`")
    nao_mapeadas = [c for c in df_raw.columns if c not in mapa_colunas.values()]
    if nao_mapeadas:
        st.markdown("**Não utilizadas:** " + ", ".join(f"`{c}`" for c in nao_mapeadas))

df = preparar_dados(df_raw, mapa_colunas)

if df.empty:
    st.error("❌ Após validação, nenhum registro válido restou.")
    st.stop()


# ---------------------------------------------------------------------------
# FILTROS LATERAIS ADICIONAIS
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
    st.sidebar.markdown("**Selecione o intervalo:**")
    col_a, col_b = st.sidebar.columns(2)
    with col_a:
        ini_input = st.date_input("Data Início",
                                  value=(hoje - timedelta(days=7)).date())
    with col_b:
        fim_input = st.date_input("Data Fim", value=hoje.date())
    if ini_input > fim_input:
        st.sidebar.error("⚠️ Data início > data fim.")
        st.stop()
    data_inicio = pd.Timestamp(ini_input)
    data_fim = pd.Timestamp(fim_input) + timedelta(days=1) - timedelta(seconds=1)

# Filtro por Analista (se existir)
analista_sel = None
if "Analista" in df.columns:
    analistas = sorted(df["Analista"].dropna().unique().tolist())
    analista_sel = st.sidebar.multiselect(
        "🧑‍💻 Filtrar por Analista",
        options=analistas,
        default=[],
        help="Deixe vazio para incluir todos.",
    )

st.sidebar.markdown("---")
st.sidebar.caption(
    f"🗓️ De **{data_inicio.strftime('%d/%m/%Y %H:%M')}** "
    f"até **{data_fim.strftime('%d/%m/%Y %H:%M')}**"
)


# ---------------------------------------------------------------------------
# FILTRAGEM
# ---------------------------------------------------------------------------
df_filtrado = df[(df["Data"] >= data_inicio) & (df["Data"] <= data_fim)].copy()

if analista_sel:
    df_filtrado = df_filtrado[df_filtrado["Analista"].isin(analista_sel)]

if df_filtrado.empty:
    st.warning("⚠️ Nenhum chamado no período/filtros selecionados.")
    st.stop()


# ---------------------------------------------------------------------------
# KPIs ADAPTATIVOS
# ---------------------------------------------------------------------------
st.markdown("### 📊 Indicadores do Período")

total = len(df_filtrado)
kpis = [("📞 Total de Chamados", f"{total:,}".replace(",", "."))]

# Taxa de Resolução N1 (via Status ou via Resolvido_N1)
if "Resolvido_N1" in df_filtrado.columns:
    resolvidos = int((df_filtrado["Resolvido_N1"] == "Sim").sum())
    taxa = (resolvidos / total * 100) if total else 0.0
    kpis.append(("✅ Taxa de Resolução (FCR)", f"{taxa:.1f}%"))
elif "Status" in df_filtrado.columns:
    status_lower = df_filtrado["Status"].str.lower()
    resolvidos = int(status_lower.str.contains("resolvido", na=False).sum())
    taxa = (resolvidos / total * 100) if total else 0.0
    kpis.append(("✅ Taxa de Resolução (N1)", f"{taxa:.1f}%"))

    cancelados = int(status_lower.str.contains("cancelad", na=False).sum())
    taxa_canc = (cancelados / total * 100) if total else 0.0
    kpis.append(("🚫 Taxa de Cancelamento", f"{taxa_canc:.1f}%"))

# TMA
if "Tempo_Atendimento_Min" in df_filtrado.columns:
    tma = df_filtrado["Tempo_Atendimento_Min"].mean()
    kpis.append(("🎧 TMA Médio", f"{tma:.1f} min" if pd.notna(tma) else "N/A"))
elif "TMA_Calculado_Min" in df_filtrado.columns:
    tma = df_filtrado["TMA_Calculado_Min"].mean()
    if pd.notna(tma):
        kpis.append(("🎧 TMA Médio (Data Encerr. - Data)", f"{tma:.1f} min"))

# TME
if "Tempo_Espera_Min" in df_filtrado.columns:
    tme = df_filtrado["Tempo_Espera_Min"].mean()
    kpis.append(("⏱️ TME Médio", f"{tme:.1f} min" if pd.notna(tme) else "N/A"))

# Analistas ativos
if "Analista" in df_filtrado.columns:
    n = df_filtrado["Analista"].nunique()
    kpis.append(("👥 Analistas Ativos", f"{n}"))

# Chamados em aberto (heurística por Status)
if "Status" in df_filtrado.columns:
    em_aberto = df_filtrado["Status"].str.lower().str.contains(
        r"novo|aberto|em atendimento|aguardando|pendente", regex=True, na=False
    ).sum()
    kpis.append(("📬 Em Aberto / Andamento", f"{int(em_aberto)}"))

# Renderiza em grade
COLS = 4
for i in range(0, len(kpis), COLS):
    linha = kpis[i:i + COLS]
    cols = st.columns(len(linha))
    for col, (titulo, valor) in zip(cols, linha):
        col.metric(titulo, valor)

st.markdown("---")


# ---------------------------------------------------------------------------
# GRÁFICOS
# ---------------------------------------------------------------------------
col_g1, col_g2 = st.columns(2)

with col_g1:
    if "Motivo" in df_filtrado.columns:
        st.markdown("#### 🔎 Top Assuntos / Motivos")
        mc = df_filtrado["Motivo"].value_counts().head(15).reset_index()
        mc.columns = ["Motivo", "Quantidade"]
        fig = px.bar(mc, x="Quantidade", y="Motivo", orientation="h",
                     text="Quantidade", color="Quantidade",
                     color_continuous_scale="Blues")
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Nº de Chamados", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("ℹ️ Sem coluna de Motivo/Assunto.")

with col_g2:
    rotulo = period_option.split(" ")[0]
    st.markdown(f"#### 📈 Evolução de Volume — Visão {rotulo}")
    st.plotly_chart(
        build_time_evolution_chart(df_filtrado, period_option, data_inicio, data_fim),
        use_container_width=True,
    )


col_g3, col_g4 = st.columns(2)

with col_g3:
    if "Status" in df_filtrado.columns:
        st.markdown("#### 📌 Distribuição por Status")
        sc = df_filtrado["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Quantidade"]
        fig = px.pie(sc, names="Status", values="Quantidade", hole=0.5)
        fig.update_layout(margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True)

with col_g4:
    if "Analista" in df_filtrado.columns:
        st.markdown("#### 🧑‍💻 Top Analistas por Volume")
        ac = df_filtrado["Analista"].value_counts().head(15).reset_index()
        ac.columns = ["Analista", "Quantidade"]
        fig = px.bar(ac, x="Quantidade", y="Analista", orientation="h",
                     text="Quantidade", color="Quantidade",
                     color_continuous_scale="Greens")
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Nº de Chamados", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)


# --- Heatmap Hora x Dia da Semana -----------------------------------------
st.markdown("#### 🔥 Heatmap de Volume — Hora × Dia da Semana")
st.caption(
    "Útil para dimensionar escala e identificar picos de demanda por "
    "faixa horária e dia da semana."
)
st.plotly_chart(build_heatmap_hora_dia(df_filtrado), use_container_width=True)


# --- Ranking detalhado de Analistas (tabela) -------------------------------
if "Analista" in df_filtrado.columns and "Status" in df_filtrado.columns:
    st.markdown("#### 🏆 Ranking Detalhado de Analistas")
    df_rank = df_filtrado.copy()
    df_rank["Resolvido"] = df_rank["Status"].str.lower().str.contains(
        "resolvido", na=False
    )
    df_rank["Cancelado"] = df_rank["Status"].str.lower().str.contains(
        "cancelad", na=False
    )

    aggs = {
        "Total": ("Status", "size"),
        "Resolvidos": ("Resolvido", "sum"),
        "Cancelados": ("Cancelado", "sum"),
    }
    if "TMA_Calculado_Min" in df_rank.columns:
        aggs["TMA Médio (min)"] = ("TMA_Calculado_Min", "mean")

    ranking = df_rank.groupby("Analista").agg(**aggs).reset_index()
    ranking["Taxa Resolução (%)"] = (
        ranking["Resolvidos"] / ranking["Total"] * 100
    ).round(1)
    ranking["Taxa Cancelamento (%)"] = (
        ranking["Cancelados"] / ranking["Total"] * 100
    ).round(1)
    if "TMA Médio (min)" in ranking.columns:
        ranking["TMA Médio (min)"] = ranking["TMA Médio (min)"].round(1)

    ranking = ranking.sort_values("Total", ascending=False)
    st.dataframe(ranking, use_container_width=True, hide_index=True)

st.markdown("---")


# ---------------------------------------------------------------------------
# AUDITORIA
# ---------------------------------------------------------------------------
st.markdown("### 🗂️ Auditoria de Chamados (Dados Brutos)")
st.caption(
    "Tabela detalhada dos registros filtrados. Use o botão abaixo para "
    "exportar em CSV (compatível com Excel)."
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
