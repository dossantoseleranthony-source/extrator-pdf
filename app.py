import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import zipfile
from datetime import datetime

import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_bytes

# =========================
# CONFIG
# =========================
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

MAX_SIZE_MB = 50
MAX_HISTORICO = 5  # usado no histórico abaixo

st.set_page_config(page_title="Extrator Inteligente", layout="wide")

# =========================
# CSS — escopado por classe customizada para não vazar em outros botões
# =========================
st.markdown("""
<style>
/* Botão principal de ação */
div[data-testid="stButton"] > button {
    background-color: #4CAF50;
    color: white;
    border-radius: 8px;
    border: none;
    font-weight: bold;
    transition: opacity 0.2s;
}
div[data-testid="stButton"] > button:hover {
    opacity: 0.85;
}

/* Botões de download */
div[data-testid="stDownloadButton"] > button {
    background-color: #008CBA;
    color: white;
    border-radius: 8px;
    border: none;
    transition: opacity 0.2s;
}
div[data-testid="stDownloadButton"] > button:hover {
    opacity: 0.85;
}
</style>
""", unsafe_allow_html=True)

# =========================
# SESSION STATE
# =========================
if "historico" not in st.session_state:
    st.session_state.historico = []

if "resultados" not in st.session_state:
    st.session_state.resultados = []

# =========================
# LIMPEZA
# =========================

def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Limpa nomes de colunas e garante unicidade com contador por nome."""
    seen: dict[str, int] = {}
    novas: list[str] = []

    for i, col in enumerate(df.columns):
        nome = str(col).strip().replace("\n", " ")
        if not nome or nome.lower() == "none" or len(nome) > 40:
            nome = f"col_{i}"

        count = seen.get(nome, 0)
        seen[nome] = count + 1
        novas.append(f"{nome}_{count}" if count else nome)

    df.columns = novas
    return df


def limpar_df(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas/colunas vazias, preenche NaN e filtra linhas de paginação."""
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
    df = df.fillna("")
    mascara_pag = df.apply(
        lambda row: row.astype(str).str.contains(r"\bpág", case=False).any(), axis=1
    )
    return df[~mascara_pag].reset_index(drop=True)


# =========================
# PDF
# =========================

def processar_pdf(pdf_bytes: bytes) -> tuple[list[pd.DataFrame], list[str]]:
    tabelas: list[pd.DataFrame] = []
    logs: list[str] = []

    config_linhas = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_y_tolerance": 15,
    }
    config_texto = {"vertical_strategy": "text", "horizontal_strategy": "text"}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pagina in enumerate(pdf.pages):
            tb = pagina.extract_tables(config_linhas) or pagina.extract_tables(config_texto) or []
            logs.append(f"Página {i+1}: {len(tb)} tabela(s) encontrada(s)")

            for tabela in tb:
                df = pd.DataFrame(tabela)
                if df.empty or df.shape[1] <= 1:
                    continue

                df = limpar_df(df)
                if df.empty:
                    continue

                primeira = df.iloc[0].astype(str)
                if primeira.str.len().mean() > 2:
                    df.columns = primeira
                    df = df.iloc[1:].reset_index(drop=True)

                df = normalizar_colunas(df)
                tabelas.append(df)

    return tabelas, logs


# =========================
# OCR AVANÇADO (OpenCV + Tesseract)
# =========================

def extrair_tabela_super(img: Image.Image) -> tuple[pd.DataFrame | None, list[str]]:
    logs: list[str] = []

    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_RGB2GRAY)  # PIL é RGB — não BGR

    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        15, 4
    )

    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, h_kernel)

    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, v_kernel)

    table_mask = cv2.add(horizontal, vertical)

    contours, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    boxes = sorted([cv2.boundingRect(c) for c in contours], key=lambda b: (b[1], b[0]))

    rows: list[list] = []
    current_row: list = []
    last_y = -1

    for (x, y, w, h) in boxes:
        if last_y == -1:
            last_y = y

        threshold = h * 0.5  # threshold dinâmico — adapta ao tamanho da célula

        if abs(y - last_y) > threshold:
            if current_row:
                rows.append(current_row)
            current_row = []
            last_y = y

        current_row.append((x, y, w, h))

    if current_row:
        rows.append(current_row)

    tabela: list[list[str]] = []

    for row in rows:
        row_sorted = sorted(row, key=lambda b: b[0])
        linha: list[str] = []
        for (x, y, w, h) in row_sorted:
            cell = img_cv[y:y+h, x:x+w]
            text = pytesseract.image_to_string(cell, config="--oem 3 --psm 6 -l por")
            linha.append(text.strip())
        tabela.append(linha)

    if tabela:
        df = normalizar_colunas(pd.DataFrame(tabela))
        logs.append("Tabela detectada via OpenCV + OCR")
        return df, logs

    return None, logs


# =========================
# OCR SIMPLES (fallback)
# =========================

def extrair_tabela_ocr(img: Image.Image) -> tuple[pd.DataFrame | None, list[str]]:
    logs: list[str] = []

    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_RGB2GRAY)

    data = (
        pytesseract
        .image_to_data(gray, output_type=pytesseract.Output.DATAFRAME)
        .dropna()
    )

    tabela: list[list[str]] = []
    max_cols = 0

    for _, linha in data.groupby("line_num"):
        row = linha.sort_values("left")["text"].tolist()
        if row:
            tabela.append(row)
            max_cols = max(max_cols, len(row))

    if tabela:
        tabela_pad = [r + [""] * (max_cols - len(r)) for r in tabela]
        df = normalizar_colunas(pd.DataFrame(tabela_pad))
        logs.append("OCR simples aplicado")
        return df, logs

    return None, logs


# =========================
# PIPELINE PRINCIPAL
# =========================

@st.cache_data(show_spinner=False)
def processar_arquivo(bytes_file: bytes, tipo: str) -> tuple[list[pd.DataFrame], list[str]]:
    """Extrai tabelas de PDF ou imagem. Cacheado por conteúdo do arquivo."""
    tabelas: list[pd.DataFrame] = []
    logs: list[str] = []

    if "pdf" in tipo:
        tb, log_pdf = processar_pdf(bytes_file)
        tabelas.extend(tb)
        logs.extend(log_pdf)

        if not tabelas:
            logs.append("Nenhuma tabela encontrada — iniciando OCR")
            imagens = convert_from_bytes(bytes_file)

            for i, img in enumerate(imagens):
                df, log = extrair_tabela_super(img)
                if df is None:
                    df, log2 = extrair_tabela_ocr(img)
                    log.extend(log2)
                logs.extend([f"Página {i+1}: {l}" for l in log])
                if df is not None:
                    tabelas.append(df)
    else:
        img = Image.open(io.BytesIO(bytes_file))
        df, log = extrair_tabela_super(img)
        if df is None:
            df, log2 = extrair_tabela_ocr(img)
            log.extend(log2)
        logs.extend(log)
        if df is not None:
            tabelas.append(df)

    return tabelas, logs


# =========================
# EXCEL
# =========================

def gerar_excel(tabelas: list[pd.DataFrame], nome_base: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for i, df in enumerate(tabelas):
            nome = re.sub(r"[\\/*?:\[\]]", "", f"{nome_base[:15]}_Tb{i+1}")[:31]
            df.to_excel(writer, index=False, sheet_name=nome)
    return buffer.getvalue()


# =========================
# UI PRINCIPAL
# =========================

st.title("📄 Extrator Inteligente")
st.caption("PDFs e imagens → Excel automaticamente")

arquivos = st.file_uploader(
    "Envie arquivos",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if arquivos:
    st.info(f"{len(arquivos)} arquivo(s) carregado(s)")

    if st.button("🚀 Processar"):
        progresso = st.progress(0)
        status = st.empty()
        total = len(arquivos)

        st.session_state.resultados = []

        for i, arquivo in enumerate(arquivos):
            status.text(f"🔍 Processando {arquivo.name}")

            if arquivo.size > MAX_SIZE_MB * 1024 * 1024:
                st.error(f"{arquivo.name} excede o limite de {MAX_SIZE_MB}MB")
                progresso.progress((i + 1) / total)
                continue

            try:
                tabelas, logs = processar_arquivo(arquivo.getvalue(), arquivo.type)
            except Exception as e:
                st.error(f"Erro em {arquivo.name}: {e}")
                progresso.progress((i + 1) / total)
                continue

            if tabelas:
                status.text("📊 Gerando Excel...")
                excel = gerar_excel(tabelas, arquivo.name)

                st.session_state.resultados.append({
                    "nome": arquivo.name,
                    "qtd": len(tabelas),
                    "excel": excel,
                    "tabelas": tabelas,
                    "logs": logs,
                    "data": datetime.now().strftime("%d/%m %H:%M"),
                })

                # Histórico limitado a MAX_HISTORICO entradas
                st.session_state.historico.append({
                    "nome": arquivo.name,
                    "data": datetime.now().strftime("%d/%m %H:%M"),
                    "qtd": len(tabelas),
                    "excel": excel,
                })
                st.session_state.historico = st.session_state.historico[-MAX_HISTORICO:]

            else:
                st.warning(f"{arquivo.name}: nenhuma tabela encontrada")

            progresso.progress((i + 1) / total)

        status.success("✅ Finalizado!")

# =========================
# RESULTADOS — fora do bloco if st.button para não somírem em reruns
# =========================

if st.session_state.resultados:
    st.markdown("## 📊 Resultados")

    # FIX: ZIP dentro do bloco de resultados, com key única
    if len(st.session_state.resultados) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for item in st.session_state.resultados:
                nome_limpo = item["nome"].rsplit(".", 1)[0]  # FIX: rsplit para nomes com pontos
                zf.writestr(f"{nome_limpo}.xlsx", item["excel"])

        st.download_button(
            "📦 Baixar TODOS (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="tabelas_extraidas.zip",
            key="dl_zip"
        )

    for i, item in enumerate(st.session_state.resultados):
        st.markdown(f"**{item['nome']} — {item['qtd']} tabela(s)**")

        st.download_button(
            "⬇️ Baixar Excel",
            data=item["excel"],
            file_name=f"{item['nome'].rsplit('.', 1)[0]}.xlsx",  # FIX: rsplit
            key=f"dl_{i}_{item['nome']}"  # FIX: key única por item — evita DuplicateWidgetID
        )

        with st.expander("Preview"):
            for df in item["tabelas"]:
                st.dataframe(df, use_container_width=True)  # FIX: restaurado use_container_width

        with st.expander("Logs"):
            for log in item["logs"]:
                st.text(log)

        st.divider()

# =========================
# HISTÓRICO
# =========================

if st.session_state.historico:
    st.markdown("## 📜 Histórico")

    for i, item in enumerate(reversed(st.session_state.historico)):
        st.markdown(f"**📄 {item['nome']}**")
        st.caption(f"{item['data']} • {item['qtd']} tabelas")

        st.download_button(
            "⬇️ Baixar novamente",
            data=item["excel"],
            file_name=f"{item['nome'].rsplit('.', 1)[0]}.xlsx",
            key=f"hist_{i}_{item['nome']}"  # key única no histórico também
        )

        st.divider()
