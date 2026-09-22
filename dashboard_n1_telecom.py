"""
Dashboard de Suporte N1 - Telecom (v6.3 - Final)
================================================
Correções desta versão:
  - _fig_to_png() agora reporta a causa REAL do erro (não só "instale kaleido").
  - render_exportacao_dashboard() cai automaticamente para HTML quando o PNG falha.
  - Nova _criar_dashboard_html() como fallback que só depende do plotly.
  - Base64 do logo/mascote movidos para constantes no topo — se vazias,
    o código segue funcionando (só não mostra a imagem).

Como executar:
    pip install -r requirements.txt
    streamlit run dashboard_n1_telecom.py
"""

import base64
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
    from PIL import Image
    _TEM_PIL = True
except ImportError:
    _TEM_PIL = False

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
# IDENTIDADE VISUAL DA MARCA (VaraNet)
# ===========================================================================
# ⚠️ Cole aqui os base64 do logo e do mascote que você já usa.
# Se deixar vazio, o banner renderiza apenas com o texto — nada quebra.
LOGO_BASE64 = ""      # <-- cole aqui o base64 do logo
MASCOTE_BASE64 = ""   # <-- cole aqui o base64 do mascote

# Cores extraídas do logo e do mascote.
AZUL_MARCA = "#00688A"
AZUL_MARCA_CLARO = "#0092C4"
TEAL_LOGO = "#11ADA9"
ROSA_LOGO = "#E8057E"
LARANJA_LOGO = "#EC7612"


def _img_tag(b64: str, alt: str, css_class: str = "") -> str:
    """Gera a tag <img> apenas se houver base64 válido."""
    if not b64 or b64.startswith("COLE"):
        return ""
    classe = f' class="{css_class}"' if css_class else ""
    return f'<img src="data:image/png;base64,{b64}" alt="{alt}"{classe} />'


def aplicar_estilo_marca():
    """Injeta CSS para aplicar o azul da marca em toda a interface."""
    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{ border-right: 3px solid {AZUL_MARCA}; }}
        div[data-baseweb="tab-highlight"],
        .stButton>button[kind="primary"],
        .stDownloadButton>button {{
            background-color: {AZUL_MARCA} !important;
            border-color: {AZUL_MARCA} !important;
            color: #FFFFFF !important;
        }}
        .stDownloadButton>button:hover {{
            background-color: {AZUL_MARCA_CLARO} !important;
        }}
        div[data-testid="stMetric"] {{
            background: linear-gradient(180deg, rgba(0,104,138,0.06) 0%,
                                        rgba(0,104,138,0.00) 100%);
            border-top: 3px solid {AZUL_MARCA};
            border-radius: 8px;
            padding: 12px 10px 6px 10px;
        }}
        div[data-testid="stMetricValue"] {{ color: {AZUL_MARCA}; }}
        h1, h2, h3 {{ color: {AZUL_MARCA}; }}
        .banner-marca {{
            background: linear-gradient(90deg, {AZUL_MARCA} 0%,
                                        {TEAL_LOGO} 100%);
            border-radius: 12px;
            padding: 18px 26px;
            display: flex;
            align-items: center;
            gap: 18px;
            margin-bottom: 18px;
            box-shadow: 0 4px 14px rgba(0,104,138,0.25);
        }}
        .banner-marca img {{ height: 46px; }}
        .banner-marca .banner-textos h1 {{
            color: #FFFFFF !important;
            margin: 0;
            font-size: 1.6rem;
            line-height: 1.2;
        }}
        .banner-marca .banner-textos p {{
            color: #EAF6F8;
            margin: 2px 0 0 0;
            font-size: 0.92rem;
        }}
        .mascote-sidebar {{ text-align: center; margin-top: 18px; }}
        .mascote-sidebar img {{ max-width: 150px; }}
        .mascote-sidebar p {{
            font-size: 0.82rem;
            color: {AZUL_MARCA};
            font-weight: 600;
            margin-top: 4px;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_banner_marca():
    """Banner de topo com o logo da VaraNet."""
    logo_html = _img_tag(LOGO_BASE64, "Logo VaraNet")
    st.markdown(
        f"""
        <div class="banner-marca">
            {logo_html}
            <div class="banner-textos">
                <h1>📡 Dashboard de Suporte N1 - Telecom</h1>
                <p>Painel automático de indicadores, evolução temporal e
                auditoria — Squad N1</p>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_mascote_sidebar(mensagem="Time N1 mandando bem! 🎉"):
    """Mostra o mascote da marca no rodapé da barra lateral."""
    mascote_html = _img_tag(MASCOTE_BASE64, "Mascote VaraNet")
    if not mascote_html:
        return
    st.sidebar.markdown(
        f"""
        <div class="mascote-sidebar">
            {mascote_html}
            <p>{mensagem}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


aplicar_estilo_marca()


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
        io.BytesIO(raw), sep=sep, encoding=enc, header=None,
        dtype=str, keep_default_na=False, na_values=[""],
        engine="python", on_bad_lines="skip",
    )
    if df_bruto.empty:
        return df_bruto

    idx_cab = _detectar_linha_cabecalho(df_bruto)
    df = pd.read_csv(
        io.BytesIO(raw), sep=sep, encoding=enc, header=idx_cab,
        dtype=str, keep_default_na=False, na_values=[""],
        engine="python", on_bad_lines="skip",
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
    "Protocolo": [r"^protocolo", r"^protocol$", r"^numero_protocolo"],
    "Solicitante": [
        r"^solicitante", r"^cliente", r"^customer", r"^assinante", r"^conta",
    ],
    "Motivo": [
        r"^assunto", r"^motivo", r"^categoria", r"^tipo", r"^topico",
        r"^classificacao", r"^problema",
    ],
    "Setor": [r"^setor", r"^departamento", r"^fila", r"^grupo"],
    "Analista": [
        r"^analista", r"^atendente", r"^operador", r"^agente",
        r"^responsavel", r"^tecnico", r"^owner", r"^assigned",
    ],
    "Status": [r"^status$", r"^situacao", r"^estado", r"^state"],
    "Urgencia": [r"^urgencia", r"^prioridade", r"^priority", r"^severidade"],
    "Interno": [r"^interno", r"^internal", r"^is_internal"],
    "Resolvido_N1": [r"^resolvido", r"^resolvido_n1", r"^fcr", r"^solucionado"],
    "Tempo_Espera_Min": [
        r"^tempo_espera", r"^tme", r"^espera", r"^wait_time", r"^fila",
    ],
    "Tempo_Atendimento_Min": [
        r"^tempo_atendimento", r"^tma", r"^duracao", r"^aht", r"^handle_time",
    ],
}

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


# ===========================================================================
# CLASSIFICAÇÃO N1 x N2
# ===========================================================================
N2_PADROES_DEFAULT = [r"provision", r"bater.*cto", r"\bcto\b"]


def eh_provavel_n2(valor) -> bool:
    texto_norm = _normalizar(valor)
    return any(re.search(padrao, texto_norm) for padrao in N2_PADROES_DEFAULT)


def sugerir_tags_n2(valores) -> list:
    vistos = []
    for v in valores:
        if v not in vistos and eh_provavel_n2(v):
            vistos.append(v)
    return vistos


def selecionar_tags_n2(tags_unicas, chave_widget, local=st.sidebar):
    sugestao = sugerir_tags_n2(tags_unicas)
    with local.expander("🛠️ Quais tags são N2 (Campo)?", expanded=False):
        st.caption(
            "Provisionamento, Bater CTO e outras tarefas de campo entram "
            "aqui. Ajuste a lista se necessário — o restante é tratado "
            "como N1 (suporte remoto)."
        )
        selecionadas = st.multiselect(
            "Tags/Motivos de N2",
            options=tags_unicas,
            default=sugestao,
            key=chave_widget,
            label_visibility="collapsed",
        )
    return selecionadas


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
    colunas_norm = {col: _normalizar(col) for col in df.columns}
    achados = {"tag": None, "qtd": None, "pct": None}
    for col, cn in colunas_norm.items():
        if achados["tag"] is None and any(t in cn for t in CAMPOS_AGREGADOS_TAG):
            achados["tag"] = col
        elif achados["qtd"] is None and any(t in cn for t in CAMPOS_AGREGADOS_QTD):
            achados["qtd"] = col
        elif achados["pct"] is None and any(t in cn for t in CAMPOS_AGREGADOS_PCT):
            achados["pct"] = col
    if achados["pct"] is None:
        for col in df.columns:
            if str(col).strip() == "%":
                achados["pct"] = col
                break
    achados["_eh_agregado"] = bool(achados["tag"] and achados["qtd"])
    return achados


# ===========================================================================
# LEITURA EM CACHE
# ===========================================================================
@st.cache_data(show_spinner="Lendo planilha...")
def load_data(caminho_ou_arquivo, origem, mtime, tamanho):
    return ler_planilha(caminho_ou_arquivo)


# ===========================================================================
# HELPERS DE NORMALIZAÇÃO
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
    r"novo|aberto|aguardando|em atendimento|andamento|pendente": "Em Aberto",
    r"escalad|transferid|crrc":                          "Escalado",
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
# PREPARAÇÃO DOS DADOS
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
# GRÁFICOS
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
        xaxis_title="Hora", yaxis_title="",
        margin=dict(l=10, r=10, t=30, b=10), height=320,
    )
    return fig


# ===========================================================================
# EXPORTAÇÃO DE DASHBOARD EM IMAGEM / HTML
# ===========================================================================
def _fig_to_png(fig, width=1800, height=900, scale=2):
    """
    Converte um gráfico Plotly em PNG. Se kaleido não estiver disponível
    ou falhar, levanta RuntimeError com a causa real.
    """
    try:
        import kaleido  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "O pacote `kaleido` não está instalado no ambiente. "
            "Adicione `kaleido==0.2.1` ao requirements.txt e faça o "
            "deploy novamente."
        ) from exc

    try:
        return fig.to_image(format="png", width=width,
                            height=height, scale=scale)
    except Exception as exc:
        raise RuntimeError(
            f"Falha ao exportar PNG via kaleido: "
            f"{type(exc).__name__}: {exc}. "
            "Verifique se `kaleido==0.2.1` está no requirements.txt. "
            "Se o problema persistir, use o botão 'Exportar em HTML' abaixo."
        ) from exc


def _criar_dashboard_png(df_base, os_df=None, titulo="Dashboard de Suporte N1"):
    """Monta uma imagem única, em alta resolução, com KPIs e gráficos."""
    if not _TEM_PIL:
        raise RuntimeError(
            "Instale Pillow para gerar a imagem: pip install pillow"
        )

    from PIL import Image, ImageDraw, ImageFont

    if "Data" not in df_base.columns and "Quantidade" in df_base.columns:
        total = int(pd.to_numeric(df_base["Quantidade"],
                                  errors="coerce").fillna(0).sum())
    else:
        total = len(df_base)

    resolvidos = cancelados = 0
    if "Status" in df_base.columns:
        status = df_base["Status"].astype(str).str.lower()
        resolvidos = int(status.str.contains(
            "resolvido|conclu|fechad|finalizad|encerrad",
            regex=True, na=False).sum())
        cancelados = int(status.str.contains(
            "cancel|ausent|desist", regex=True, na=False).sum())

    os_total = len(os_df) if os_df is not None else 0

    cards = [
        ("TOTAL DE ATENDIMENTOS", f"{total:,}".replace(",", ".")),
        ("RESOLVIDOS", f"{resolvidos:,}".replace(",", ".")),
        ("CANCELADOS", f"{cancelados:,}".replace(",", ".")),
        ("ABERTURAS DE O.S.", f"{os_total:,}".replace(",", ".")),
    ]

    figs = []
    if "Data" in df_base.columns:
        fig_volume = build_time_evolution_chart(df_base)
        fig_volume.update_layout(title="Evolução de Volume", font=dict(size=16))
        figs.append(("volume", fig_volume, 1800, 800))

    if "Motivo" in df_base.columns:
        mc = df_base["Motivo"].value_counts().head(12).reset_index()
        mc.columns = ["Motivo", "Quantidade"]
        fig_mot = px.bar(mc, x="Quantidade", y="Motivo", orientation="h",
                         text="Quantidade", color="Quantidade",
                         color_continuous_scale="Blues")
        fig_mot.update_layout(
            title="Atendimentos por Motivo",
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=55, b=10),
        )
        fig_mot.update_traces(textposition="outside")
        figs.append(("motivos", fig_mot, 1800, 900))

    if os_df is not None and not os_df.empty and "Motivo" in os_df.columns:
        oc = os_df["Motivo"].value_counts().head(12).reset_index()
        oc.columns = ["Motivo", "Quantidade"]
        fig_os = px.bar(oc, x="Quantidade", y="Motivo", orientation="h",
                        text="Quantidade", color="Quantidade",
                        color_continuous_scale="Oranges")
        fig_os.update_layout(
            title="Aberturas de O.S. por Motivo",
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=55, b=10),
        )
        fig_os.update_traces(textposition="outside")
        figs.append(("os", fig_os, 1800, 900))

    rendered = []
    for _, fig, w, h in figs:
        rendered.append(
            Image.open(io.BytesIO(_fig_to_png(fig, w, h, 2))).convert("RGB")
        )

    margin = 60
    card_h = 190
    gap = 35
    title_h = 150
    chart_gap = 45
    total_h = (title_h + card_h + chart_gap
               + sum(im.height for im in rendered) + gap * (len(rendered) + 2))
    canvas_w = max([im.width for im in rendered] + [1800]) + margin * 2
    canvas = Image.new("RGB", (canvas_w, total_h), "white")
    draw = ImageDraw.Draw(canvas)

    try:
        font_title = ImageFont.truetype("DejaVuSans-Bold.ttf", 54)
        font_card = ImageFont.truetype("DejaVuSans-Bold.ttf", 24)
        font_value = ImageFont.truetype("DejaVuSans-Bold.ttf", 48)
    except Exception:
        font_title = font_card = font_value = ImageFont.load_default()

    y = 45
    draw.text((margin, y), titulo, fill="#00688A", font=font_title)
    y += title_h

    card_w = (canvas_w - margin * 2 - gap * 3) // 4
    for idx, (label, value) in enumerate(cards):
        x = margin + idx * (card_w + gap)
        draw.rounded_rectangle((x, y, x + card_w, y + card_h), radius=24,
                               fill="#F5F8FA", outline="#D8E3E8", width=3)
        draw.text((x + 25, y + 25), label, fill="#4B5B63", font=font_card)
        draw.text((x + 25, y + 82), value, fill="#00688A", font=font_value)

    y += card_h + chart_gap
    for im in rendered:
        x = (canvas_w - im.width) // 2
        canvas.paste(im, (x, y))
        y += im.height + gap

    out = io.BytesIO()
    canvas.save(out, format="PNG", optimize=True)
    return out.getvalue()


def _criar_dashboard_html(df_base, os_df=None,
                          titulo="Dashboard de Suporte N1"):
    """
    Gera um HTML único com KPIs (em cartões) e todos os gráficos Plotly.
    Não depende de kaleido nem Pillow — só do plotly.
    """
    import plotly.io as pio

    if "Data" not in df_base.columns and "Quantidade" in df_base.columns:
        total = int(pd.to_numeric(df_base["Quantidade"],
                                  errors="coerce").fillna(0).sum())
    else:
        total = len(df_base)

    resolvidos = cancelados = 0
    if "Status" in df_base.columns:
        status = df_base["Status"].astype(str).str.lower()
        resolvidos = int(status.str.contains(
            "resolvido|conclu|fechad|finalizad|encerrad",
            regex=True, na=False).sum())
        cancelados = int(status.str.contains(
            "cancel|ausent|desist", regex=True, na=False).sum())

    os_total = len(os_df) if os_df is not None else 0

    def _card(label, valor):
        return (
            f'<div style="flex:1;min-width:200px;background:#F5F8FA;'
            f'border:2px solid #D8E3E8;border-radius:16px;padding:18px;">'
            f'<div style="color:#4B5B63;font-size:13px;font-weight:700;'
            f'letter-spacing:.5px;">{label}</div>'
            f'<div style="color:#00688A;font-size:38px;font-weight:700;">'
            f'{valor}</div></div>'
        )

    cards_html = (
        '<div style="display:flex;gap:18px;flex-wrap:wrap;'
        'margin:18px 0 26px 0;">'
        + _card("TOTAL DE ATENDIMENTOS", f"{total:,}".replace(",", "."))
        + _card("RESOLVIDOS", f"{resolvidos:,}".replace(",", "."))
        + _card("CANCELADOS", f"{cancelados:,}".replace(",", "."))
        + _card("ABERTURAS DE O.S.", f"{os_total:,}".replace(",", "."))
        + '</div>'
    )

    blocos = []
    if "Data" in df_base.columns:
        fig_vol = build_time_evolution_chart(df_base)
        fig_vol.update_layout(title="Evolução de Volume",
                              height=520, font=dict(size=16))
        blocos.append(pio.to_html(fig_vol, include_plotlyjs=False,
                                  full_html=False))

    if "Motivo" in df_base.columns:
        mc = df_base["Motivo"].value_counts().head(12).reset_index()
        mc.columns = ["Motivo", "Quantidade"]
        fig_mot = px.bar(mc, x="Quantidade", y="Motivo", orientation="h",
                         text="Quantidade", color="Quantidade",
                         color_continuous_scale="Blues")
        fig_mot.update_layout(
            title="Atendimentos por Motivo", height=620,
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=55, b=10),
        )
        fig_mot.update_traces(textposition="outside")
        blocos.append(pio.to_html(fig_mot, include_plotlyjs=False,
                                  full_html=False))

    if os_df is not None and not os_df.empty and "Motivo" in os_df.columns:
        oc = os_df["Motivo"].value_counts().head(12).reset_index()
        oc.columns = ["Motivo", "Quantidade"]
        fig_os = px.bar(oc, x="Quantidade", y="Motivo", orientation="h",
                        text="Quantidade", color="Quantidade",
                        color_continuous_scale="Oranges")
        fig_os.update_layout(
            title="Aberturas de O.S. por Motivo", height=620,
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=55, b=10),
        )
        fig_os.update_traces(textposition="outside")
        blocos.append(pio.to_html(fig_os, include_plotlyjs=False,
                                  full_html=False))

    corpo_graficos = "\n".join(
        f'<div style="background:white;border-radius:14px;padding:18px;'
        f'margin-bottom:22px;box-shadow:0 2px 10px rgba(0,0,0,0.06);">'
        f'{b}</div>'
        for b in blocos
    )

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8" />
<title>{titulo}</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI",
          Roboto, sans-serif; background:#EEF3F6; margin:0; padding:32px; }}
  h1 {{ color:#00688A; margin:0 0 4px 0; font-size:28px; }}
  .sub {{ color:#4B5B63; margin:0 0 22px 0; font-size:14px; }}
  .container {{ max-width:1400px; margin:0 auto; }}
</style>
</head>
<body>
<div class="container">
  <h1>📡 {titulo}</h1>
  <p class="sub">Squad N1 — Telecom · Gerado em
  {datetime.now():%d/%m/%Y %H:%M}</p>
  {cards_html}
  {corpo_graficos}
</div>
</body>
</html>"""

    return html.encode("utf-8")


def render_exportacao_dashboard(df_base, os_df=None):
    st.markdown("### 🖼️ Exportação do Dashboard")
    st.caption(
        "Gera uma imagem única em alta resolução com o volume de "
        "atendimentos, principais gráficos e uma área específica para "
        "as aberturas de O.S."
    )

    # --- Tentativa de PNG ---------------------------------------------------
    try:
        png = _criar_dashboard_png(df_base, os_df=os_df)
        st.download_button(
            "⬇️ Exportar Dashboard em PNG — Alta Qualidade",
            data=png,
            file_name="dashboard_suporte_n1_alta_qualidade.png",
            mime="image/png",
            use_container_width=True,
        )
        return
    except RuntimeError as exc:
        st.warning(
            "⚠️ **Não foi possível gerar o PNG.** "
            "Usando o fallback em HTML, que você pode abrir no navegador "
            "e salvar como PDF ou imprimir como imagem.\n\n"
            f"**Motivo técnico:** {exc}"
        )

    # --- Fallback HTML ------------------------------------------------------
    try:
        html = _criar_dashboard_html(df_base, os_df=os_df)
        st.download_button(
            "⬇️ Exportar Dashboard em HTML (abre no navegador)",
            data=html,
            file_name="dashboard_suporte_n1.html",
            mime="text/html",
            use_container_width=True,
        )
        st.caption(
            "💡 Depois de baixar o HTML: abra no Chrome/Edge → "
            "**Ctrl+P** → **Salvar como PDF** (ou \"Microsoft Print to PDF\"). "
            "O resultado sai em alta qualidade, igual ao PNG."
        )
    except Exception as exc:
        st.error(f"❌ Não foi possível gerar nem o HTML de exportação: {exc}")


# ===========================================================================
# DASHBOARD DE O.S.
# ===========================================================================
def render_dashboard_os(df_base):
    if "Motivo" not in df_base.columns:
        st.info("Não foi encontrada a coluna Motivo/Assunto para montar "
                "o painel de O.S.")
        return

    motivo = df_base["Motivo"].astype(str)
    mascara_os = motivo.str.contains(
        r"\bO\.?\s*S\.?\b|ordem\s+de\s+servi|link\s*loss|abertura\s+de\s+os",
        case=False, regex=True, na=False
    )
    os_df = df_base.loc[mascara_os].copy()

    total_os = len(os_df)
    total_geral = len(df_base)
    pct_os = (total_os / total_geral * 100) if total_geral else 0.0

    st.markdown("### 🛠️ Dashboard de Abertura de O.S.")
    st.caption(
        "Área reservada para acompanhar exclusivamente os atendimentos "
        "relacionados à abertura, informação ou tratamento de Ordem de "
        "Serviço."
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("🧾 O.S. identificadas", f"{total_os:,}".replace(",", "."))
    c2.metric("% dos atendimentos", f"{pct_os:.1f}%")
    if not os_df.empty and "Analista" in os_df.columns:
        c3.metric("👤 Analistas envolvidos",
                  f"{os_df['Analista'].nunique():,}".replace(",", "."))
    else:
        c3.metric("👤 Analistas envolvidos", "-")

    if os_df.empty:
        st.info("Nenhum atendimento classificado como O.S. foi encontrado.")
        return

    col1, col2 = st.columns(2)
    with col1:
        oc = os_df["Motivo"].value_counts().head(15).reset_index()
        oc.columns = ["Motivo", "Quantidade"]
        fig = px.bar(oc, x="Quantidade", y="Motivo", orientation="h",
                     text="Quantidade", color="Quantidade",
                     color_continuous_scale="Oranges")
        fig.update_layout(
            title="Tipos de O.S.",
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            margin=dict(l=10, r=10, t=55, b=10), height=500,
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        if "Analista" in os_df.columns:
            ac = os_df["Analista"].value_counts().head(15).reset_index()
            ac.columns = ["Analista", "Quantidade"]
            fig = px.bar(ac, x="Quantidade", y="Analista", orientation="h",
                         text="Quantidade", color="Quantidade",
                         color_continuous_scale="Oranges")
            fig.update_layout(
                title="O.S. por Analista",
                yaxis={"categoryorder": "total ascending"},
                showlegend=False, coloraxis_showscale=False,
                margin=dict(l=10, r=10, t=55, b=10), height=500,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

    if "Status" in os_df.columns:
        st.markdown("#### 📌 Status das O.S.")
        sc = os_df["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Quantidade"]
        fig = px.pie(sc, names="Status", values="Quantidade", hole=0.5)
        fig.update_layout(margin=dict(l=10, r=10, t=30, b=10), height=420)
        fig.update_traces(textposition="inside", textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)

    if "Data" in os_df.columns:
        st.markdown("#### 📈 Evolução das Aberturas de O.S.")
        os_time = os_df.copy()
        os_time["Dia"] = os_time["Data"].dt.normalize()
        serie = os_time.groupby("Dia").size().reset_index(name="O.S.")
        fig = px.bar(serie, x="Dia", y="O.S.", text="O.S.")
        fig.update_layout(
            xaxis_title="Data", yaxis_title="Quantidade de O.S.",
            margin=dict(l=10, r=10, t=30, b=10), height=420,
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### 🗂️ Registros relacionados a O.S.")
    st.dataframe(
        os_df.sort_values("Data", ascending=False)
        if "Data" in os_df.columns else os_df,
        use_container_width=True, hide_index=True,
    )

    csv_os = os_df.to_csv(index=False, sep=";").encode("utf-8-sig")
    st.download_button(
        "⬇️ Exportar dados de O.S. (CSV)",
        data=csv_os,
        file_name="abertura_os_n1.csv",
        mime="text/csv",
    )


# ===========================================================================
# MODO AGREGADO (Tag, Quantidade, %)
# ===========================================================================
def render_modo_agregado(df: pd.DataFrame, agregados: dict):
    render_banner_marca()
    st.caption(
        "O arquivo carregado é um **resumo** (Tag, Quantidade, %). "
        "Exibindo painel agregado. Se você quiser os KPIs de tickets, "
        "carregue o export detalhado (com coluna `Data`)."
    )

    col_tag = agregados["tag"]
    col_qtd = agregados["qtd"]

    df_view = df.copy()
    df_view[col_qtd] = pd.to_numeric(df_view[col_qtd], errors="coerce")
    df_view = df_view.dropna(subset=[col_qtd]).sort_values(col_qtd,
                                                          ascending=False)

    tags_unicas = df_view[col_tag].dropna().astype(str).unique().tolist()
    tags_n2 = selecionar_tags_n2(tags_unicas, chave_widget="tags_n2_agregado")
    df_view["Nível"] = df_view[col_tag].apply(
        lambda t: "N2 (Campo)" if t in tags_n2 else "N1 (Remoto)"
    )

    df_n1 = df_view[df_view["Nível"] == "N1 (Remoto)"]
    df_n2 = df_view[df_view["Nível"] == "N2 (Campo)"]

    total = int(df_view[col_qtd].sum())
    total_n1 = int(df_n1[col_qtd].sum())
    total_n2 = int(df_n2[col_qtd].sum())
    pct_n2 = (total_n2 / total * 100) if total else 0.0
    top_tag = df_view.iloc[0][col_tag] if len(df_view) else "-"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Total Geral", f"{total:,}".replace(",", "."))
    c2.metric("🧑‍💻 Total N1 (Remoto)", f"{total_n1:,}".replace(",", "."))
    c3.metric("🛠️ Total N2 (Campo)", f"{total_n2:,}".replace(",", "."))
    c4.metric("📊 % N2 do Total", f"{pct_n2:.1f}%")
    st.caption(f"🥇 Tag líder geral: **{top_tag}**")
    st.markdown("---")

    def _bloco_tag(df_nivel, titulo, cor_escala):
        if df_nivel.empty:
            st.info(f"Nenhuma tag classificada como {titulo} neste arquivo.")
            return

        col_b, col_p = st.columns([2, 1])
        with col_b:
            st.markdown(f"#### 🏷️ Distribuição — {titulo}")
            top_n = df_nivel.head(20)
            fig = px.bar(
                top_n, x=col_qtd, y=col_tag, orientation="h",
                text=col_qtd, color=col_qtd,
                color_continuous_scale=cor_escala,
            )
            fig.update_layout(
                yaxis={"categoryorder": "total ascending"},
                showlegend=False, coloraxis_showscale=False,
                xaxis_title="Quantidade", yaxis_title="",
                margin=dict(l=10, r=10, t=30, b=10), height=480,
            )
            fig.update_traces(textposition="outside")
            st.plotly_chart(fig, use_container_width=True)

        with col_p:
            st.markdown("#### 🥧 Top 10 (proporção)")
            top10 = df_nivel.head(10)
            fig = px.pie(top10, names=col_tag, values=col_qtd, hole=0.5)
            fig.update_traces(textposition="inside", textinfo="percent",
                              insidetextfont=dict(size=11))
            fig.update_layout(
                margin=dict(l=10, r=10, t=10, b=10), height=520,
                legend=dict(orientation="h", yanchor="top", y=-0.05,
                            xanchor="center", x=0.5, font=dict(size=10)),
                uniformtext_minsize=8, uniformtext_mode="hide",
            )
            st.plotly_chart(fig, use_container_width=True)

        st.dataframe(
            df_nivel[[col_tag, col_qtd]].rename(
                columns={col_tag: "Tag", col_qtd: "Quantidade"}
            ),
            use_container_width=True, hide_index=True,
        )

    aba_geral, aba_n1, aba_n2 = st.tabs(
        ["📊 Visão Geral", "🧑‍💻 N1 - Suporte Remoto", "🛠️ N2 - Campo"]
    )
    with aba_geral:
        _bloco_tag(df_view, "Todas as Tags", "Blues")
    with aba_n1:
        _bloco_tag(df_n1, "N1 (Suporte Remoto)", "Blues")
    with aba_n2:
        _bloco_tag(df_n2, "N2 (Campo)", "Oranges")

    st.markdown("---")
    st.markdown("### 🗂️ Tabela completa (com classificação N1/N2)")
    st.dataframe(df_view, use_container_width=True, hide_index=True)

    st.markdown("---")
    render_exportacao_dashboard(
        df_view.rename(columns={col_tag: "Motivo"}), os_df=None
    )

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
render_banner_marca()
st.caption(
    "💡 Dica: depois que o painel carregar, é só tirar um **print da tela** "
    "e mandar no grupo — o banner acima já sai formatado para compartilhar."
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

render_mascote_sidebar()

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

if usar_upload and uploaded_file is not None:
    nome_fonte = uploaded_file.name
    origem_fonte = "upload manual"
else:
    nome_fonte = ARQUIVO_AUTO
    origem_fonte = "arquivo automático"

st.info(f"📄 **Lendo:** `{nome_fonte}`  \n🔎 **Origem:** {origem_fonte}")


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

agregados = detectar_modo_agregado(df_raw)
if agregados["_eh_agregado"]:
    render_modo_agregado(df_raw, agregados)
    st.stop()

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
    kpis.append(("👥 Analistas Ativos", f"{df['Analista'].nunique()}"))

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
# GRÁFICOS
# ===========================================================================
if "Motivo" in df.columns:
    st.markdown("### 🔎 Distribuição por Motivo (N1 x N2)")
    motivos_unicos = df["Motivo"].dropna().astype(str).unique().tolist()
    tags_n2_detalhado = selecionar_tags_n2(
        motivos_unicos, chave_widget="tags_n2_detalhado"
    )
    df["Nível_Suporte"] = df["Motivo"].apply(
        lambda m: "N2 (Campo)" if m in tags_n2_detalhado else "N1 (Remoto)"
    )

    total_geral = len(df)
    total_n1_det = int((df["Nível_Suporte"] == "N1 (Remoto)").sum())
    total_n2_det = int((df["Nível_Suporte"] == "N2 (Campo)").sum())
    pct_n2_det = (total_n2_det / total_geral * 100) if total_geral else 0.0

    m1, m2, m3 = st.columns(3)
    m1.metric("🧑‍💻 Chamados N1 (Remoto)",
              f"{total_n1_det:,}".replace(",", "."))
    m2.metric("🛠️ Chamados N2 (Campo)",
              f"{total_n2_det:,}".replace(",", "."))
    m3.metric("📊 % N2 do Total", f"{pct_n2_det:.1f}%")

    def _grafico_motivo(df_nivel, cor_escala):
        if df_nivel.empty:
            st.info("Nenhum chamado nesta categoria no período.")
            return
        mc = df_nivel["Motivo"].value_counts().head(15).reset_index()
        mc.columns = ["Motivo", "Quantidade"]
        fig = px.bar(mc, x="Quantidade", y="Motivo", orientation="h",
                     text="Quantidade", color="Quantidade",
                     color_continuous_scale=cor_escala)
        fig.update_layout(
            yaxis={"categoryorder": "total ascending"},
            showlegend=False, coloraxis_showscale=False,
            xaxis_title="Nº de Chamados", yaxis_title="",
            margin=dict(l=10, r=10, t=30, b=10),
        )
        fig.update_traces(textposition="outside")
        st.plotly_chart(fig, use_container_width=True)

    aba_geral_m, aba_n1_m, aba_n2_m = st.tabs(
        ["📊 Todos os Motivos", "🧑‍💻 N1 - Suporte Remoto", "🛠️ N2 - Campo"]
    )
    with aba_geral_m:
        _grafico_motivo(df, "Blues")
    with aba_n1_m:
        _grafico_motivo(df[df["Nível_Suporte"] == "N1 (Remoto)"], "Blues")
    with aba_n2_m:
        _grafico_motivo(df[df["Nível_Suporte"] == "N2 (Campo)"], "Oranges")

    st.markdown("---")
else:
    st.info("ℹ️ Sem coluna de Motivo/Assunto.")


col_g1, col_g2 = st.columns(2)

with col_g1:
    st.markdown("#### 📈 Evolução de Volume na Base")
    st.plotly_chart(build_time_evolution_chart(df), use_container_width=True)

with col_g2:
    if "Status" in df.columns:
        st.markdown("#### 📌 Distribuição por Status")
        sc = df["Status"].value_counts().reset_index()
        sc.columns = ["Status", "Quantidade"]
        fig = px.pie(
            sc, names="Status", values="Quantidade", hole=0.5,
            color_discrete_sequence=[AZUL_MARCA, TEAL_LOGO,
                                     LARANJA_LOGO, ROSA_LOGO],
        )
        fig.update_traces(textposition="inside", textinfo="percent",
                          insidetextfont=dict(size=11))
        fig.update_layout(
            margin=dict(l=10, r=10, t=10, b=10), height=420,
            legend=dict(orientation="h", yanchor="top", y=-0.05,
                        xanchor="center", x=0.5, font=dict(size=10)),
            uniformtext_minsize=8, uniformtext_mode="hide",
        )
        st.plotly_chart(fig, use_container_width=True)


if "Analista" in df.columns:
    st.markdown("#### 🧑‍💻 Top Analistas por Volume")
    ac = df["Analista"].value_counts().head(15).reset_index()
    ac.columns = ["Analista", "Quantidade"]
    fig = px.bar(ac, x="Quantidade", y="Analista", orientation="h",
                 text="Quantidade", color="Quantidade",
                 color_continuous_scale="Teal")
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


# ===========================================================================
# ÁREA DE O.S.
# ===========================================================================
render_dashboard_os(df)

st.markdown("---")

# O.S. para exportação
os_df_export = None
if "Motivo" in df.columns:
    os_df_export = df.loc[
        df["Motivo"].astype(str).str.contains(
            r"\bO\.?\s*S\.?\b|ordem\s+de\s+servi|link\s*loss|abertura\s+de\s+os",
            case=False, regex=True, na=False
        )
    ].copy()

render_exportacao_dashboard(df, os_df=os_df_export)


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
