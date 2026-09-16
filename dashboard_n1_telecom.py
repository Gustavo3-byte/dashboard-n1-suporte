"""
Dashboard de Suporte N1 - Telecom
==================================
Ferramenta web de relatórios automáticos e dashboard interativo para
squads de atendimento Nível 1 (N1) de Telecom.

Como executar:
    pip install streamlit pandas plotly openpyxl
    streamlit run dashboard_n1_telecom.py


import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime, timedelta


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO GERAL DA PÁGINA
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Dashboard N1 - Telecom",
    page_icon="📡",
    layout="wide",
)


# ---------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Processando arquivo carregado...")
def load_data(file):
    """Lê o arquivo enviado (csv ou xlsx) e retorna um DataFrame bruto."""
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


def preparar_dados(df_raw):
    """Valida colunas, converte tipos e limpa a base."""
    df = df_raw.copy()

    # Boas práticas de manipulação temporal: conversão robusta para datetime,
    # evitando que valores inválidos quebrem os filtros por índice de data.
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)

    registros_invalidos = int(df["Data"].isna().sum())
    if registros_invalidos > 0:
        st.warning(
            f"⚠️ {registros_invalidos} registro(s) com data/hora inválida foram "
            "descartados automaticamente da análise."
        )
        df = df.dropna(subset=["Data"])

    # Garante que os campos numéricos realmente sejam numéricos
    for col in ["Tempo_Espera_Min", "Tempo_Atendimento_Min"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # Padroniza texto de FCR ('sim' / 'SIM' / 'Sim' -> 'Sim')
    df["Resolvido_N1"] = df["Resolvido_N1"].astype(str).str.strip().str.capitalize()
    df["Motivo"] = df["Motivo"].astype(str).str.strip()
    df["Status"] = df["Status"].astype(str).str.strip()

    df = df.sort_values("Data").reset_index(drop=True)
    return df


def build_time_evolution_chart(df_periodo, period_option, data_inicio, data_fim):
    """
    Constrói o gráfico de evolução temporal adaptado ao escopo escolhido:
      - Diário       -> agregação HORA A HORA
      - Semanal      -> agregação DIA A DIA
      - Mensal       -> agregação POR SEMANA
      - Personalizado-> escolhe a granularidade de acordo com o tamanho do intervalo
    df_local = df_periodo.copy()
    delta_dias = (data_fim.normalize() - data_inicio.normalize()).days

    def agrega_por_hora(inicio, fim):
        df_local["Agrupador"] = df_local["Data"].dt.floor("H")
        faixa_completa = pd.date_range(start=inicio.floor("H"), end=fim.floor("H"), freq="H")
        contagem = (
            df_local.groupby("Agrupador").size().reindex(faixa_completa, fill_value=0).reset_index()
        )
        contagem.columns = ["Periodo", "Chamados"]
        fig = px.line(contagem, x="Periodo", y="Chamados", markers=True)
        fig.update_xaxes(title_text="Hora do Dia", tickformat="%H:%M")
        return fig

    def agrega_por_dia(inicio, fim):
        df_local["Agrupador"] = df_local["Data"].dt.normalize()
        faixa_completa = pd.date_range(start=inicio.normalize(), end=fim.normalize(), freq="D")
        contagem = (
            df_local.groupby("Agrupador").size().reindex(faixa_completa, fill_value=0).reset_index()
        )
        contagem.columns = ["Periodo", "Chamados"]
        fig = px.line(contagem, x="Periodo", y="Chamados", markers=True)
        fig.update_xaxes(title_text="Dia", tickformat="%d/%m")
        return fig

    def agrega_por_semana():
        df_local["Agrupador"] = df_local["Data"].dt.to_period("W").apply(lambda r: r.start_time)
        contagem = df_local.groupby("Agrupador").size().reset_index()
        contagem.columns = ["Periodo", "Chamados"]
        contagem = contagem.sort_values("Periodo")
        contagem["Rotulo"] = contagem["Periodo"].apply(lambda d: f"Semana de {d.strftime('%d/%m')}")
        fig = px.line(contagem, x="Rotulo", y="Chamados", markers=True)
        fig.update_xaxes(title_text="Semana")
        return fig

    if period_option == "Diário (Hoje)":
        fig = agrega_por_hora(data_inicio, data_fim)
    elif period_option == "Semanal (Últimos 7 dias)":
        fig = agrega_por_dia(data_inicio, data_fim)
    elif period_option == "Mensal (Últimos 30 dias)":
        fig = agrega_por_semana()
    else:
        # Período Personalizado: granularidade adaptativa
        if delta_dias <= 1:
            fig = agrega_por_hora(data_inicio, data_fim)
        elif delta_dias <= 31:
            fig = agrega_por_dia(data_inicio, data_fim)
        else:
            fig = agrega_por_semana()

    fig.update_traces(line=dict(width=3), fill="tozeroy", fillcolor="rgba(0,123,255,0.12)")
    fig.update_layout(
        yaxis_title="Nº de Chamados",
        hovermode="x unified",
        margin=dict(l=10, r=10, t=30, b=10),
    )
    return fig


# ---------------------------------------------------------------------------
# CABEÇALHO
# ---------------------------------------------------------------------------
st.title("📡 Dashboard de Suporte N1 - Telecom")
st.caption("Painel automático de indicadores, evolução temporal e auditoria para squads de atendimento Nível 1.")


# ---------------------------------------------------------------------------
# BARRA LATERAL - UPLOAD
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ Configurações do Relatório")
uploaded_file = st.sidebar.file_uploader(
    "📂 Carregue a base de atendimentos (.csv ou .xlsx)",
    type=["csv", "xlsx"],
)

# ---------------------------------------------------------------------------
# ESTADO SEM ARQUIVO CARREGADO
# ---------------------------------------------------------------------------
if uploaded_file is None:
    st.info(
        "👋 **Bem-vindo(a)!** Para começar, carregue o arquivo de atendimentos "
        "(.csv ou .xlsx) na barra lateral à esquerda."
    )
    st.markdown("#### 📋 Estrutura de colunas esperada no arquivo:")
    st.dataframe(
        pd.DataFrame(
            {
                "Coluna": REQUIRED_COLUMNS,
                "Descrição": [
                    "Data e hora do chamado (timestamp)",
                    "ID único do ticket",
                    "Motivo do contato (Sem Sinal, Lentidão, Financeiro, Configuração de Modem...)",
                    "Status do chamado (Aberto, Resolvido, Escalado)",
                    "Tempo de espera em minutos",
                    "Tempo de atendimento em minutos",
                    "'Sim' ou 'Não' - se o N1 resolveu o chamado (base do cálculo de FCR)",
                ],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    st.stop()


# ---------------------------------------------------------------------------
# LEITURA E VALIDAÇÃO DO ARQUIVO
# ---------------------------------------------------------------------------
try:
    df_raw = load_data(uploaded_file)
except Exception as erro:
    st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes técnicos: {erro}")
    st.stop()

colunas_faltantes = [c for c in REQUIRED_COLUMNS if c not in df_raw.columns]
if colunas_faltantes:
    st.error(
        "❌ O arquivo carregado não contém todas as colunas obrigatórias.\n\n"
        f"**Colunas ausentes:** {', '.join(colunas_faltantes)}\n\n"
        f"**Colunas esperadas:** {', '.join(REQUIRED_COLUMNS)}"
    )
    st.stop()

df = preparar_dados(df_raw)

if df.empty:
    st.error("❌ Após a validação dos dados, nenhum registro válido restou para análise. Verifique o arquivo enviado.")
    st.stop()


# ---------------------------------------------------------------------------
# BARRA LATERAL - SELETOR DE PERÍODO
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

hoje = pd.Timestamp(datetime.now().date())  # meia-noite do dia atual

if period_option == "Diário (Hoje)":
    data_inicio = hoje
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)

elif period_option == "Semanal (Últimos 7 dias)":
    data_inicio = hoje - timedelta(days=6)
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)

elif period_option == "Mensal (Últimos 30 dias)":
    data_inicio = hoje - timedelta(days=29)
    data_fim = hoje + timedelta(days=1) - timedelta(seconds=1)

else:  # Período Personalizado
    st.sidebar.markdown("**Selecione o intervalo desejado:**")
    col_data_a, col_data_b = st.sidebar.columns(2)
    with col_data_a:
        data_inicio_input = st.date_input("Data Início", value=(hoje - timedelta(days=7)).date())
    with col_data_b:
        data_fim_input = st.date_input("Data Fim", value=hoje.date())

    if data_inicio_input > data_fim_input:
        st.sidebar.error("⚠️ A data de início não pode ser posterior à data de fim.")
        st.stop()

    data_inicio = pd.Timestamp(data_inicio_input)
    data_fim = pd.Timestamp(data_fim_input) + timedelta(days=1) - timedelta(seconds=1)

st.sidebar.markdown("---")
st.sidebar.caption(
    f"🗓️ Exibindo dados de **{data_inicio.strftime('%d/%m/%Y %H:%M')}** "
    f"até **{data_fim.strftime('%d/%m/%Y %H:%M')}**"
)


# ---------------------------------------------------------------------------
# FILTRAGEM DA BASE PELO PERÍODO SELECIONADO
# ---------------------------------------------------------------------------
df_filtrado = df[(df["Data"] >= data_inicio) & (df["Data"] <= data_fim)].copy()

if df_filtrado.empty:
    st.warning(
        "⚠️ Nenhum chamado foi encontrado para o período selecionado. "
        "Ajuste o filtro de período na barra lateral ou verifique se a base "
        "carregada contém registros nesta janela de tempo."
    )
    st.stop()


# ---------------------------------------------------------------------------
# BLOCO DE MÉTRICAS (KPIs)
# ---------------------------------------------------------------------------
st.markdown("### 📊 Indicadores do Período")

total_chamados = len(df_filtrado)
resolvidos_n1 = int((df_filtrado["Resolvido_N1"] == "Sim").sum())
taxa_fcr = (resolvidos_n1 / total_chamados * 100) if total_chamados > 0 else 0.0
tme_medio = df_filtrado["Tempo_Espera_Min"].mean()
tma_medio = df_filtrado["Tempo_Atendimento_Min"].mean()

col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
col_kpi1.metric("📞 Total de Chamados", f"{total_chamados:,}".replace(",", "."))
col_kpi2.metric("✅ Taxa de FCR (N1)", f"{taxa_fcr:.1f}%")
col_kpi3.metric("⏱️ TME Médio", f"{tme_medio:.1f} min" if pd.notna(tme_medio) else "N/A")
col_kpi4.metric("🎧 TMA Médio", f"{tma_medio:.1f} min" if pd.notna(tma_medio) else "N/A")

st.markdown("---")


# ---------------------------------------------------------------------------
# GRÁFICOS INTERATIVOS
# ---------------------------------------------------------------------------
col_graf1, col_graf2 = st.columns(2)

with col_graf1:
    st.markdown("#### 🔎 Principais Motivos de Contato")
    motivo_counts = df_filtrado["Motivo"].value_counts().reset_index()
    motivo_counts.columns = ["Motivo", "Quantidade"]

    fig_motivos = px.bar(
        motivo_counts,
        x="Quantidade",
        y="Motivo",
        orientation="h",
        text="Quantidade",
        color="Quantidade",
        color_continuous_scale="Blues",
    )
    fig_motivos.update_layout(
        yaxis={"categoryorder": "total ascending"},
        showlegend=False,
        coloraxis_showscale=False,
        xaxis_title="Nº de Chamados",
        yaxis_title="",
        margin=dict(l=10, r=10, t=30, b=10),
    )
    fig_motivos.update_traces(textposition="outside")
    st.plotly_chart(fig_motivos, use_container_width=True)

with col_graf2:
    rotulo_periodo = period_option.split(" ")[0]
    st.markdown(f"#### 📈 Evolução de Volume — Visão {rotulo_periodo}")
    fig_evolucao = build_time_evolution_chart(df_filtrado, period_option, data_inicio, data_fim)
    st.plotly_chart(fig_evolucao, use_container_width=True)

st.markdown("---")


# ---------------------------------------------------------------------------
# TABELA DE AUDITORIA
# ---------------------------------------------------------------------------
st.markdown("### 🗂️ Auditoria de Chamados (Dados Brutos)")
st.caption(
    "Utilize a tabela abaixo para conferência detalhada dos registros do período "
    "filtrado. Os dados podem ser exportados para planilha através do botão abaixo."
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
