import io
import chardet  # pip install chardet  (opcional; se não quiser, uso fallback)

# Ordem de tentativas de encoding — cobre praticamente todos os exports BR
ENCODINGS_TENTADOS = ["utf-8-sig", "utf-8", "latin-1", "cp1252"]

# Separadores comuns em exports de plataformas de ticket
SEPARADORES_TENTADOS = [";", ",", "\t", "|"]


def _detectar_encoding(raw_bytes: bytes) -> str:
    """Tenta adivinhar o encoding; cai em latin-1 se tudo falhar."""
    try:
        det = chardet.detect(raw_bytes[:200_000])
        if det and det.get("encoding") and det.get("confidence", 0) > 0.7:
            return det["encoding"]
    except Exception:
        pass
    return "utf-8-sig"


def _ler_csv_inteligente(source) -> pd.DataFrame:
    """
    Lê CSV tentando combinações de encoding x separador até uma funcionar
    com pelo menos 2 colunas reconhecíveis. Funciona tanto para upload
    (BytesIO) quanto para caminho em disco.
    """
    # Normaliza para bytes
    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, str):
            raw = raw.encode("utf-8", errors="ignore")
    else:
        with open(source, "rb") as f:
            raw = f.read()

    encodings = [_detectar_encoding(raw)] + ENCODINGS_TENTADOS
    vistos = set()
    ultimo_erro = None

    for enc in encodings:
        if enc in vistos:
            continue
        vistos.add(enc)
        for sep in SEPARADORES_TENTADOS:
            try:
                df = pd.read_csv(
                    io.BytesIO(raw),
                    sep=sep,
                    encoding=enc,
                    dtype=str,           # lê tudo como texto; convertemos depois
                    keep_default_na=False,
                    na_values=[""],
                    engine="c",
                )
                # Considera válido se detectou pelo menos 3 colunas
                if df.shape[1] >= 3:
                    return df
            except Exception as e:
                ultimo_erro = e
                continue

    raise RuntimeError(
        f"Não foi possível interpretar o CSV. Último erro: {ultimo_erro}"
    )


@st.cache_data(show_spinner="Carregando e validando dados...")
def load_data(uploaded_file, caminho_auto, mtime, tamanho):
    """
    Carrega dados de upload OU do arquivo automático.
    'mtime' e 'tamanho' entram só para invalidar o cache corretamente.
    """
    # Fonte 1: upload manual
    if uploaded_file is not None:
        nome = uploaded_file.name.lower()
        if nome.endswith(".csv"):
            return _ler_csv_inteligente(uploaded_file)
        return pd.read_excel(uploaded_file, dtype=str)

    # Fonte 2: arquivo automático
    if caminho_auto and os.path.exists(caminho_auto):
        if caminho_auto.lower().endswith((".xlsx", ".xls")):
            return pd.read_excel(caminho_auto, dtype=str)
        return _ler_csv_inteligente(caminho_auto)

    raise FileNotFoundError("Nenhuma fonte de dados disponível.")
