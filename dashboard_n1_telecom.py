"""
Dashboard de Suporte N1 - Telecom (v6.2)
=========================================
Novidades desta versão:
  - Mostra SEMPRE qual arquivo está sendo lido (nome + origem).
  - Detecta "modo agregado": se o arquivo tiver apenas colunas do tipo
    Tag/Quantidade/%, exibe um painel de barras/pizza direto, sem
    exigir coluna Data.
  - Mensagem de erro cirúrgica quando falta Data, com lista de colunas.

Como executar:
    pip install streamlit pandas plotly openpyxl chardet
    streamlit run dashboard_n1_telecom.py
"""

import csv
import io
import json
import os
import re
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import chardet
    _TEM_CHARDET = True
except ImportError:
    _TEM_CHARDET = False


# ===========================================================================
# CONFIGURAÇÃO GERAL
# ===========================================================================
st.set_page_config(
    page_title="Dashboard N1 - Telecom",
    page_icon="📡",
    layout="wide",
)

ARQUIVO_AUTO = "dados_n1_auto.csv"


# ===========================================================================
# LEITOR UNIVERSAL (embutido)
# ===========================================================================
ENCODINGS = ["utf-8-sig", "utf-8", "latin-1", "cp1252", "utf-16", "ascii"]
SEPARADORES = [";", ",", "\t", "|", ":"]
EXT_CSV = {".csv", ".tsv", ".txt"}
EXT_EXCEL = {".xlsx", ".xlsm", ".xls", ".xlsb", ".ods"}


def _achar_bytes(source) -> bytes:
    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8", errors="ignore")
        try:
            source.seek(0)
        except Exception:
            pass
        return raw
    with open(source, "rb") as f:
        return f.read()


def _achar_nome(source) -> str:
    if hasattr(source, "name"):
        return str(source.name)
    if isinstance(source, (str, os.PathLike)):
        return str(source)
    return ""


def _detectar_encoding(raw: bytes) -> str:
    if _TEM_CHARDET:
        try:
            det = chardet.detect(raw[:200_000])
            if det and det.get("encoding") and det.get("confidence", 0) > 0.7:
                return det["encoding"]
        except Exception:
            pass
    return "utf-8-sig"


def _detectar_separador(amostra: str) -> str:
    try:
        dialect = csv.Sniffer().sniff(amostra, delimiters=";,\t|:")
        return dialect.delimiter
    except Exception:
        for linha in amostra.splitlines():
            if linha.strip():
                contagens = {s: linha.count(s) for s in SEPARADORES}
                melhor = max(contagens, key=contagens.get)
                if contagens[melhor] > 0:
                    return melhor
                break
        return ";"


def _detectar_linha_cabecalho(df_bruto: pd.DataFrame) -> int:
    melhor_idx = 0
    melhor_score = -1

    for i in range(min(20, len(df_bruto))):
        linha = df_bruto.iloc[i].astype(str).tolist()
        preenchidas = [c for c in linha
                       if c and c.lower() not in ("nan", "none", "")]
        if len(preenchidas) < 2:
            continue

        nao_numericas = sum(
            1 for c in preenchidas
            if not c.replace(".", "").replace(",", "").replace("-", "")
                 .replace("/", "").replace(":", "").isdigit()
        )
        score = nao_numericas * 2 + len(preenchidas)

        if i + 1 < len(df_bruto):
            prox = df_bruto.iloc[i + 1].astype(str).tolist()
            preenchidas_prox = [c for c in prox
                                if c and c.lower() not in ("nan", "none", "")]
            if len(preenchidas_prox) >= len(preenchidas) - 1:
                score += 2

        if score > melhor_score:
            melhor_score = score
            melhor_idx = i

    return melhor_idx


def _ler_csv(raw: bytes) -> pd.DataFrame:
    enc = _detectar_encoding(raw)
    try:
        amostra = raw.decode(enc, errors="replace")[:50_000]
    except Exception:
        amostra = raw[:50_000].decode("latin-1", errors="replace")

    sep = _detectar_separador(amostra)

    df_bruto = pd.read_csv(
        io.BytesIO(raw),
        sep=sep,
        encoding=enc,
        header=None,
        dtype=str,
        keep_default_na=False,
        na_values=[""],
        engine="python",
        on_bad_lines="skip",
    )

    if df_bruto.empty:
        return df_bruto

    idx_cab = _detectar_linha_cabecalho(df_bruto)

    df = pd.read_csv(
        io.BytesIO(raw),
        sep=sep,
        encoding=enc,
        header=idx_cab,
        dtype=str,
        keep_default_na=False,
        na_values=[""],
        engine="python",
        on_bad_lines="skip",
    )

    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _ler_excel(source) -> pd.DataFrame:
    nome = _achar_nome(source).lower()
    engine = None
    if nome.endswith(".xlsb"):
        engine = "pyxlsb"
    elif nome.endswith(".ods"):
        engine = "odf"

    if hasattr(source, "read"):
        raw = _achar_bytes(source)
        xls = pd.ExcelFile(io.BytesIO(raw), engine=engine)
    else:
        xls = pd.ExcelFile(source, engine=engine)

    melhor_aba, melhor_score = None, -1
    for aba in xls.sheet_names:
        try:
            tmp = xls.parse(aba, header=None, dtype=str, nrows=50)
            linhas, colunas = tmp.shape
            score = linhas * colunas
            if score > melhor_score:
                melhor_score = score
                melhor_aba = aba
        except Exception:
            continue

    if melhor_aba is None:
        raise RuntimeError("Nenhuma aba legível no Excel.")

    df_bruto = xls.parse(melhor_aba, header=None, dtype=str)
    if df_bruto.empty:
        return df_bruto

    idx_cab = _detectar_linha_cabecalho(df_bruto)
    df = xls.parse(melhor_aba, header=idx_cab, dtype=str)
    df = df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _ler_json(raw: bytes) -> pd.DataFrame:
    texto = raw.decode(_detectar_encoding(raw), errors="replace")
    obj = json.loads(texto)

    if isinstance(obj, list):
        return pd.json_normalize(obj)
    if isinstance(obj, dict):
        for chave in ("data", "items", "result", "results",
                      "rows", "records"):
            if chave in obj:
                valor = obj[chave]
                if isinstance(valor, dict):
                    for chave2 in ("items", "rows", "records"):
                        if chave2 in valor:
                            return pd.json_normalize(valor[chave2])
                return pd.json_normalize(valor)
        return pd.json_normalize(obj)
    raise RuntimeError("Formato JSON não reconhecido.")


def ler_planilha(source) -> pd.DataFrame:
    nome = _achar_nome(source).lower()
    ext = os.path.splitext(nome)[1]

    if ext in EXT_CSV or ext == "":
        raw = _achar_bytes(source)
        if raw[:4] == b"PK\x03\x04":
            return _ler_excel(io.BytesIO(raw))
        if raw[:4] == b"\xd0\xcf\x11\xe0":
            return _ler_excel(io.BytesIO(raw))
        return _ler_csv(raw)

    if ext in EXT_EXCEL:
        return _ler_excel(source)

    if ext == ".json":
        return _ler_json(_achar_bytes(source))

    if ext == ".parquet":
        return pd.read_parquet(source)
    if ext == ".feather":
        return pd.read_feather(source)

    try:
        return _ler_csv(_achar_bytes(source))
    except Exception:
        return _ler_excel(source)


# ===========================================================================
# MAPEAMENTO DE COLUNAS
# ===========================================================================
FIELD_SYNONYMS = {
    "Data": [
        r"^data$", r"^data_", r"_data$", r"^dt", r"^abertura",
        r"^criado", r"^created", r"^timestamp", r"^inicio",
        r"^data.*hora", r"^data.*abert",
    ],
    "Data_Encerrado": [
        r"^data_encerrado", r"^data_fechamento", r"^data_conclusao",
        r"^data_finalizacao", r"^fechado_em", r"^closed_at",
        r"^encerrado", r"^data.*encerr", r"^data.*fech",
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

# Campos aceitos para o "modo agregado"
CAMPOS_AGREGADOS_TAG = ("tag", "categoria", "motivo", "assunto",
                        "tipo", "topico", "classificacao")
CAMPOS_AGREGADOS_QTD = ("quantidade", "qtd", "contagem", "total",
                        "count", "freq", "frequencia")
CAMPOS_AGREGADOS_PCT = ("%", "percentual", "porcentagem", "pct",
                        "share", "proporcao")


# ===========================================================================
# NORMALIZAÇÃO E DETECÇÃO
# ===========================================================================
def _normalizar(texto) -> str:
    if texto is None:
        return ""
    texto = str(texto)
    texto = (texto.replace("\ufeff", "")
                  .replace("\u00a0", " ")
                  .replace("\u200b", "")
                  .replace('"', "")
                  .replace("'", ""))
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = texto.strip().lower()
    texto = re.sub(r"[\s\-\.\/\(\)\[\]]+", "_", texto)
    texto = re.sub(r"_+", "_", texto)
    texto = texto.strip("_")
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


def _fallback_data(df: pd.DataFrame, mapa: dict) -> dict:
    if "Data" in mapa:
        return mapa
    for col in df.columns:
        cn = _normalizar(col)
        if any(t in cn for t in ("data", "dt", "abert", "criad")):
            if any(t in cn for t in ("encerr", "fech", "conclus")):
                continue
            mapa["Data"] = col
            return mapa
    return mapa


def detectar_modo_agregado(df: pd.DataFrame) -> dict:
    """
    Verifica se o arquivo é um resultado agregado (ex.: Tag, Quantidade, %).
    Devolve um dicionário {tag, qtd, pct} com os nomes reais ou None.
    """
    colunas_norm = {col: _normalizar(col) for col in df.columns}
    achados = {"tag": None, "qtd": None, "pct": None}

    for col, cn in colunas_norm.items():
        if achados["tag"] is None and any(t in cn for t in CAMPOS_AGREGADOS_TAG):
            achados["tag"] = col
        elif achados["qtd"] is None and any(t in cn for t in CAMPOS_AGREGADOS_QTD):
            achados["qtd"] = col
        elif achados["pct"] is None and any(t in cn for t in CAMPOS_AGREGADOS_PCT):
            achados["pct"] = col

    # Também aceita quando a coluna se chama literalmente "%"
    if achados["pct"] is None:
        for col in df.columns:
            if str(col).strip() == "%":
                achados["pct"] = col
                break

    # Considera "agregado" se tiver pelo menos a tag + quantidade
    achados["_eh_agregado"] = bool(achados["tag"] and achados["qtd"])
    return achados


# ===========================================================================
# LEITURA EM CACHE
# ===========================================================================
@st.cache_data(show_spinner="Lendo planilha...")
def load_data(caminho_ou_arquivo, origem, mtime, tamanho):
    if origem == "auto":
        return ler_planilha(caminho_ou_arquivo)
    return ler_planilha(caminho_ou_arquivo)


# ===========================================================================
# HELPERS DE NORMALIZAÇÃO (motivo, status, tempo)
# ===========================================================================
def _normalizar_motivo(s: pd.Series) -> pd.Series:
    s = (s.astype(str)
          .str.replace(r"\s+", " ", regex=True)
          .str.strip()
          .str.strip(".;,")
          .str.title())

    mapa_motivos = {
        r"Desbloqueio.*Confian":        "Desbloqueio por Confiança",
        r"Desbloqueio.*Pagamento":      "Desbloqueio Após Pagamento",
        r"Informou.*Pagamento":         "Informou Pagamento",
        r"Pagou.*(Boleto|Frente)":      "Pagou Boleto a Frente",
        r"Boleto.*Frente":              "Pagou Boleto a Frente",
        r"Envio.*Boleto":               "Envio de Boleto",
        r"^Boleto$":                    "Envio de Boleto",
        r"Troca.*Senha":                "Troca de Senha",
        r"Login.*Senha.*Central":       "Login/Senha Central do Assinante",
        r"Ordem.*Servi..Link Loss":     "OS Link Loss",
        r"Link Loss":                   "OS Link Loss",
        r"Instabilidade":               "Instabilidade",
        r"Se Ausentou|Ausentou.*Atend": "Ausentou do Atendimento",
        r"Inatividade":                 "Inatividade",
        r"Transferido.*CRRC":           "Transferido para o CRRC",
        r"Informa..o.*O\.?S":           "Informação sobre OS",
        r"O\.?S\.? (Geral|Retorno)":    "OS Geral/Retorno",
        r"Acesso Normalizado":          "Acesso Normalizado",
        r"Atualiza..o.*Rede":           "Atualização de Rede",
        r"Tomada.*Curto":               "Tomada em Curto",
        r"Troca.*Titularidade":         "Troca de Titularidade",
        r"Mudan.a.*Endere":             "Mudança de Endereço",
        r"Suporte.*Cliente":            "Suporte ao Cliente",
    }
    for padrao, rotulo in mapa_motivos.items():
        s = s.where(~s.str.contains(padrao, case=False, regex=True, na=False),
                    rotulo)
    return s


STATUS_CANONICOS = {
    r"resolv|conclu|fechad|solucion|finalizad|encerrad": "Resolvido",
    r"cancel|ausent|desist":                             "Cancelado",
    r"novo|aberto|aguardando|em atendimento|andamento|pendente":
                                                          "Em Aberto",
    r"escalad|transferid|crrc":                           "Escalado",
}


def _padronizar_status(serie: pd.Series) -> pd.Series:
    s = serie.astype(str).str.strip()
    resultado = s.copy()
    for padrao, rotulo in STATUS_CANONICOS.items():
        resultado = resultado.where(
            ~s.str.contains(padrao, case=False, regex=True, na=False),
            rotulo,
        )
    return resultado


def _converter_tempo_flexivel(serie: pd.Series) -> pd.Series:
    num = pd.to_numeric(serie, errors="coerce")
    if num.notna().mean() > 0.8:
        return num

    def _parse(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        s = str(v).strip()
        if not s:
            return None
        if ":" in s:
            partes = [int(p) for p in s.split(":") if p.strip().isdigit()]
            if len(partes) == 3:
                return partes[0] * 60 + partes[1] + partes[2] / 60
            if len(partes) == 2:
                return partes[0] + partes[1] / 60
        m = re.match(r"^(?:(\d+)\s*m(?:in)?)?\s*(?:(\d+)\s*s)?$", s, re.I)
        if m and (m.group(1) or m.group(2)):
            return int(m.group(1) or 0) + int(m.group(2) or 0) / 60
        return None

    return serie.map(_parse)


# ===========================================================================
# PREPARAÇÃO DOS DADOS (base detalhada)
# ===========================================================================
def preparar_dados(df_raw: pd.DataFrame, mapa: dict) -> pd.DataFrame:
    df = df_raw.copy()
    df = df.rename(columns={v: k for k, v in mapa.items()})

    for col_data in ("Data", "Data_Encerrado"):
        if col_data in df.columns:
            df[col_data] = pd.to_datetime(
                df[col_data], errors="coerce", dayfirst=True
            )

    if "Data" in df.columns:
        invalidas = int(df["Data"].isna().sum())
        if invalidas > 0:
            st.warning(
                f"⚠️ {invalidas} registro(s) com data inválida "
                "foram descartados."
            )
            df = df.dropna(subset=["Data"])

    chave_dedup = None
    if "Chamado" in df.columns:
        chave_dedup = "Chamado"
    elif "Protocolo" in df.columns:
        chave_dedup = "Protocolo"

    if chave_dedup:
        antes = len(df)
        df = (df.sort_values("Data", ascending=False)
                .drop_duplicates(subset=[chave_dedup], keep="first")
                .reset_index(drop=True))
        removidos = antes - len(df)
        if removidos > 0:
            st.info(
                f"🧹 {removidos} registro(s) duplicado(s) por `{chave_dedup}` "
                "removidos (mantido o mais recente)."
            )

    for col in ("Tempo_Espera_Min", "Tempo_Atendimento_Min"):
        if col in df.columns:
            df[col] = _converter_tempo_flexivel(df[col])

    if "Data_Encerrado" in df.columns and "Data" in df.columns:
        delta = (df["Data_Encerrado"] - df["Data"]).dt.total_seconds() / 60.0
        df["TMA_Calculado_Min"] = delta.where(delta > 0)

    if "Motivo" in df.columns:
        df["Motivo_Original"] = df["Motivo"]
        df["Motivo"] = _normalizar_motivo(df["Motivo"])

    if "Status" in df.columns:
        df["Status_Original"] = df["Status"]
        df["Status"] = _padronizar_status(df["Status"])

    for col in ("Analista", "Urgencia", "Setor", "Solicitante"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "Resolvido_N1" in df.columns:
        df["Resolvido_N1"] = (
            df["Resolvido_N1"].astype(str).str.strip().str.capitalize()
        )

    if "Data" in df.columns:
        df["Ano"] = df["Data"].dt.year
        df["Mes"] = df["Data"].dt.month
        df["Dia"] = df["Data"].dt.day
        df["DiaSemana"] = df["Data"].dt.weekday
        dias_pt = ["Segunda", "Terça", "Quarta", "Quinta",
                   "Sexta", "Sábado", "Domingo"]
        df["DiaSemanaNome"] = df["DiaSemana"].map(lambda i: dias_pt[i])
        df["Hora"] = df["Data"].dt.hour
        df["FimDeSemana"] = df["DiaSemana"] >= 5

        bins = [-1, 5, 11, 17, 23]
        labels = ["Madrugada (00-05)", "Manhã (06-11)",
                  "Tarde (12-17)", "Noite (18-23)"]
        df["FaixaHoraria"] = pd.cut(df["Hora"], bins=bins, labels=labels)

    antes = len(df)
    if "Chamado" in df.columns:
        df = df[df["Chamado"].notna() &
                (df["Chamado"].astype(str).str.strip() != "")]
    elif "Protocolo" in df.columns:
        df = df[df["Protocolo"].notna() &
                (df["Protocolo"].astype(str).str.strip() != "")]
    removidos = antes - len(df)
    if removidos > 0:
        st.caption(f"🧹 {removidos} linha(s) sem ID/protocolo removidas.")

    df = df.sort_values("Data").reset_index(drop=True)
    return df


# ===========================================================================
# GRÁFICOS (base detalhada)
# ===========================================================================
def build_time_evolution_chart(df_base):
    df_local = df_base.copy()
    dt_min = df_local["Data"].min()
    dt_max = df_local["Data"].max()
    delta_dias = (dt_max.normalize() - dt_min.normalize()).days

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

    if delta_dias <= 1:
        df_local["Agrupador"] = df_local["Data"].dt.floor("h")
        faixa = pd.date_range(start=dt_min.floor("h"),
                              end=dt_max.floor("h"), freq="h")
        contagem = (df_local.groupby("Agrupador").size()
                    .reindex(faixa, fill_value=0).reset_index())
        contagem.columns = ["Periodo", "Chamados"]
        return _linha(contagem, "Hora do Dia", "%H:%M")

    if delta_dias <= 31:
        df_local["Agrupador"] = df_local["Data"].dt.normalize()
        faixa = pd.date_range(start=dt_min.normalize(),
                              end=dt_max.normalize(), freq="D")
        contagem = (df_local.groupby("Agrupador").size()
                    .reindex(faixa, fill_value=0).reset_index())
        contagem.columns = ["Periodo", "Chamados"]
        return _linha(contagem, "Dia", "%d/%m")

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


def build_heatmap_hora_dia(df_base):
    dias_pt = ["Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom"]
    df_h = df_base.copy()
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


# ===========================================================================
# RENDERIZAÇÃO — MODO AGREGADO (Tag, Quantidade, %)
# ===========================================================================
def render_modo_agregado(df: pd.DataFrame, agregados: dict):
    st.title("📊 Painel Agregado - Suporte N1")
    st.caption(
        "O arquivo carregado é um **resumo** (Tag, Quantidade, %). "
        "Exibindo painel agregado. Se você quiser os KPIs de tickets, "
        "carregue o export detalhado (com coluna `Data`)."
    )

    col_tag = agregados["tag"]
    col_qtd = agregados["qtd"]
    col_pct = agregados["pct"]

    df_view = df.copy()
    df_view[col_qtd] = pd.to_numeric(df_view[col_qtd], errors="coerce")
    df_view = df_view.dropna(subset=[col_qtd]).sort_values(col_qtd,
                                                          ascending=False)

    total = int(df_view[col_qtd].sum())
    n_tags = int(df_view[col_tag].nunique())
    top_tag = df_view.iloc[0][col_tag] if len(df_view) else "-"
    top_qtd = int(df_view.iloc[0][col_qtd]) if len(df_view) else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Total", f"{total:,}".replace(",", "."))
    c2.metric("🏷️ Tags distintas", f"{n_tags}")
    c3.metric("🥇 Tag líder", str(top_tag)[:30])
    c4.metric("🔢 Volume da líder", f"{top_qtd:,}".replace(",", "."))

    st.markdown("---")

    col_b, col_p = st.columns([2, 1])

    with col_b:
        st.markdown("#### 🏷️ Distribuição por Tag")
        top_n = df_view.head(20)
        fig = px.bar(
            top_n, x=col_qtd, y=col_tag, orientation="h",
            text=col_qtd, color=col_qtd, color_continuous_scale="Blues",
        )
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Quantidade", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
            height=520,
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    with col_p:
        st.markdown("#### 🥧 Top 10 (proporção)")
        top10 = df_view.head(10)
        fig = px.pie(top10, names=col_tag, values=col_qtd, hole=0.5)
        fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=520)
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("### 🗂️ Tabela completa")
    st.dataframe(df_view, use_container_width=True, hide_index=True)

    csv_export = df_view.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        label="⬇️ Exportar CSV",
        data=csv_export,
        file_name="resumo_agregado_n1.csv",
        mime="text/csv",
    )


# ===========================================================================
# CABEÇALHO
# ===========================================================================
st.title("📡 Dashboard de Suporte N1 - Telecom")
st.caption(
    "Painel automático de indicadores, evolução temporal e auditoria "
    "para squads de atendimento Nível 1."
)


# ===========================================================================
# FONTE DE DADOS
# ===========================================================================
st.sidebar.header("⚙️ Fonte de Dados")

arquivo_auto_disponivel = os.path.exists(ARQUIVO_AUTO)
mtime_auto = tamanho_auto = 0
if arquivo_auto_disponivel:
    mtime_auto = os.path.getmtime(ARQUIVO_AUTO)
    tamanho_auto = os.path.getsize(ARQUIVO_AUTO)

usar_upload = st.sidebar.toggle(
    "📂 Enviar arquivo manualmente",
    value=not arquivo_auto_disponivel,
    help="Se ativado, o arquivo enviado sobrescreve o automático.",
)

uploaded_file = None
if usar_upload:
    uploaded_file = st.sidebar.file_uploader(
        "Selecione o arquivo",
        type=["csv", "tsv", "txt", "xlsx", "xlsm", "xls", "xlsb",
              "ods", "json", "parquet", "feather"],
    )
elif arquivo_auto_disponivel:
    mtime_dt = datetime.fromtimestamp(mtime_auto)
    st.sidebar.success(
        f"✅ Fonte automática ativa\n\n"
        f"Arquivo: `{ARQUIVO_AUTO}`\n\n"
        f"Última coleta: **{mtime_dt:%d/%m/%Y %H:%M}**"
    )
    if st.sidebar.button("🔄 Recarregar dados agora"):
        st.cache_data.clear()
        st.rerun()
else:
    st.sidebar.warning(
        "⚠️ Nenhuma fonte disponível.\n\n"
        f"Coloque um arquivo em `{ARQUIVO_AUTO}` "
        "ou ative o upload manual."
    )

if usar_upload and uploaded_file is None:
    st.info(
        "👋 Carregue uma planilha **ou** desative o upload "
        "para usar a coleta automática.\n\n"
        "**Formatos aceitos:** .csv, .tsv, .txt, .xlsx, .xlsm, .xls, "
        ".xlsb, .ods, .json, .parquet, .feather."
    )
    st.stop()

if not usar_upload and not arquivo_auto_disponivel:
    st.stop()

# Identificação clara do arquivo que está sendo lido
if usar_upload and uploaded_file is not None:
    nome_fonte = uploaded_file.name
    origem_fonte = "upload manual"
else:
    nome_fonte = ARQUIVO_AUTO
    origem_fonte = "arquivo automático"

st.info(
    f"📄 **Lendo:** `{nome_fonte}`  \n"
    f"🔎 **Origem:** {origem_fonte}"
)


# ===========================================================================
# LEITURA + DETECÇÃO
# ===========================================================================
try:
    if usar_upload:
        df_raw = load_data(uploaded_file, "upload", 0, 0)
    else:
        df_raw = load_data(ARQUIVO_AUTO, "auto", mtime_auto, tamanho_auto)
except Exception as erro:
    st.error(f"❌ Não foi possível carregar os dados. Detalhes: {erro}")
    st.stop()

if df_raw.empty:
    st.error("❌ Arquivo vazio.")
    st.stop()

# 1) Verifica se é um arquivo agregado (Tag, Quantidade, %)
agregados = detectar_modo_agregado(df_raw)
if agregados["_eh_agregado"]:
    render_modo_agregado(df_raw, agregados)
    st.stop()

# 2) Modo detalhado — exige coluna Data
mapa_colunas = detectar_colunas(df_raw)
mapa_colunas = _fallback_data(df_raw, mapa_colunas)

if "Data" not in mapa_colunas:
    st.error(
        "❌ Não foi possível identificar a coluna de **Data** e o arquivo "
        "também não parece ser um resumo agregado.\n\n"
        "**Colunas encontradas:** " +
        ", ".join(f"`{c}`" for c in df_raw.columns) +
        "\n\n**Esperado:** uma coluna como `Data`, `Data_Abertura`, "
        "`Criado_Em`, `Abertura` ou um arquivo agregado com "
        "`Tag` + `Quantidade`."
    )
    st.stop()

# ---------------------------------------------------------------- segue normal
with st.sidebar.expander("🧭 Colunas reconhecidas", expanded=True):
    for campo, col_real in mapa_colunas.items():
        st.markdown(f"- **{campo}** ← `{col_real}`")
    nao_mapeadas = [c for c in df_raw.columns if c not in mapa_colunas.values()]
    if nao_mapeadas:
        st.markdown("**Não utilizadas:** " +
                    ", ".join(f"`{c}`" for c in nao_mapeadas))

df = preparar_dados(df_raw, mapa_colunas)

if df.empty:
    st.error("❌ Após validação, nenhum registro válido restou.")
    st.stop()

dt_min = df["Data"].min()
dt_max = df["Data"].max()

st.caption(
    f"🗓️ Base carregada: **{len(df):,}** chamados, de "
    f"**{dt_min.strftime('%d/%m/%Y %H:%M')}** "
    f"até **{dt_max.strftime('%d/%m/%Y %H:%M')}**."
    .replace(",", ".")
)


# ===========================================================================
# KPIs
# ===========================================================================
st.markdown("### 📊 Indicadores da Base")

total = len(df)
kpis = [("📞 Total de Chamados", f"{total:,}".replace(",", "."))]

if "Resolvido_N1" in df.columns:
    resolvidos = int((df["Resolvido_N1"] == "Sim").sum())
    taxa = (resolvidos / total * 100) if total else 0.0
    kpis.append(("✅ Taxa de Resolução (FCR)", f"{taxa:.1f}%"))
elif "Status" in df.columns:
    status_lower = df["Status"].str.lower()
    resolvidos = int(status_lower.str.contains("resolvido", na=False).sum())
    taxa = (resolvidos / total * 100) if total else 0.0
    kpis.append(("✅ Taxa de Resolução (N1)", f"{taxa:.1f}%"))

    cancelados = int(status_lower.str.contains("cancelado", na=False).sum())
    taxa_canc = (cancelados / total * 100) if total else 0.0
    kpis.append(("🚫 Taxa de Cancelamento", f"{taxa_canc:.1f}%"))

if "Tempo_Atendimento_Min" in df.columns:
    tma = df["Tempo_Atendimento_Min"].mean()
    kpis.append(("🎧 TMA Médio", f"{tma:.1f} min" if pd.notna(tma) else "N/A"))
elif "TMA_Calculado_Min" in df.columns:
    tma = df["TMA_Calculado_Min"].mean()
    if pd.notna(tma):
        kpis.append(("🎧 TMA Médio (Encerr. − Abertura)", f"{tma:.1f} min"))

if "Tempo_Espera_Min" in df.columns:
    tme = df["Tempo_Espera_Min"].mean()
    kpis.append(("⏱️ TME Médio", f"{tme:.1f} min" if pd.notna(tme) else "N/A"))

if "Analista" in df.columns:
    n = df["Analista"].nunique()
    kpis.append(("👥 Analistas Ativos", f"{n}"))

if "Status" in df.columns:
    em_aberto = df["Status"].str.lower().str.contains(
        r"em aberto|novo|aberto|aguardando|pendente",
        regex=True, na=False
    ).sum()
    kpis.append(("📬 Em Aberto / Andamento", f"{int(em_aberto)}"))

COLS = 4
for i in range(0, len(kpis), COLS):
    linha = kpis[i:i + COLS]
    cols = st.columns(len(linha))
    for col, (titulo, valor) in zip(cols, linha):
        col.metric(titulo, valor)

st.markdown("---")


# ===========================================================================
# GRÁFICOS (base detalhada)
# ===========================================================================
col_g1, col_g2 = st.columns(2)

with col_g1:
    if "Motivo" in df.columns:
        st.markdown("#### 🔎 Top Assuntos / Motivos")
        mc = df["Motivo"].value_counts().head(15).reset_index()
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
    st.markdown("#### 📈 Evolução de Volume na Base")
    st.plotly_chart(build_time_evolution_chart(df), use_container_width=True)


col_g3, col_g4 = st.columns(2)

with col_g3:
    if "Status" in df.columns:
        st.markdown("#### 📌 Distribuição por Status")
        sc = df["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Quantidade"]
        fig = px.pie(sc, names="Status", values="Quantidade", hole=0.5)
        fig.update_layout(margin=dict(l=10, r=10, t=30, b=10))
        st.plotly_chart(fig, use_container_width=True)

with col_g4:
    if "Analista" in df.columns:
        st.markdown("#### 🧑‍💻 Top Analistas por Volume")
        ac = df["Analista"].value_counts().head(15).reset_index()
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


st.markdown("#### 🔥 Heatmap de Volume — Hora × Dia da Semana")
st.caption(
    "Útil para dimensionar escala e identificar picos de demanda por "
    "faixa horária e dia da semana."
)
st.plotly_chart(build_heatmap_hora_dia(df), use_container_width=True)


if "Analista" in df.columns and "Status" in df.columns:
    st.markdown("#### 🏆 Ranking Detalhado de Analistas")
    df_rank = df.copy()
    df_rank["Resolvido"] = df_rank["Status"].str.lower().str.contains(
        "resolvido", na=False
    )
    df_rank["Cancelado"] = df_rank["Status"].str.lower().str.contains(
        "cancelado", na=False
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


# ===========================================================================
# AUDITORIA
# ===========================================================================
st.markdown("### 🗂️ Auditoria de Chamados (Dados Brutos)")
st.caption(
    "Tabela detalhada dos registros. Use o botão abaixo para "
    "exportar em CSV (compatível com Excel)."
)
df_exibicao = df.sort_values("Data", ascending=False).reset_index(drop=True)
st.dataframe(df_exibicao, use_container_width=True, height=420)

csv_export = df_exibicao.to_csv(index=False, sep=";").encode("utf-8-sig")
st.download_button(
    label="⬇️ Exportar CSV",
    data=csv_export,
    file_name=f"auditoria_n1_{dt_min.date()}_a_{dt_max.date()}.csv",
    mime="text/csv",
)
