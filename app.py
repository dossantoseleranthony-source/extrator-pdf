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
MAX_HISTORICO = 5

st.set_page_config(page_title="Extrator Inteligente PRO", layout="wide")

# =========================
# SESSION STATE
# =========================
if "historico" not in st.session_state:
    st.session_state.historico = []

if "resultados" not in st.session_state:
    st.session_state.resultados = []

# =========================
# FUNÇÕES DE LIMPEZA
# =========================

def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpa nomes de colunas e garante unicidade.
    Usa contador por nome — sem índice global como sufixo.
    """
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
    # FIX v1: filtro de paginação que havia sido removido nesta versão
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
            # Tenta extração por linhas primeiro, depois por texto
            tb = pagina.extract_tables(config_linhas) or pagina.extract_tables(config_texto) or []
            logs.append(f"Página {i+1}: {len(tb)} tabela(s)")

            for tabela in tb:
                df = pd.DataFrame(tabela)
                if df.empty or df.shape[1] <= 1:
                    continue

                df = limpar_df(df)
                if df.empty:
                    continue

                # Promove primeira linha a cabeçalho se parecer com títulos
                primeira = df.iloc[0].astype(str)
                if primeira.str.len().mean() > 2:
                    df.columns = primeira
                    df = df.iloc[1:].reset_index(drop=True)

                # Normalização chamada UMA vez só aqui
                df = normalizar_colunas(df)
                tabelas.append(df)

    return tabelas, logs


# =========================
# OCR — detecção de tabela por estrutura (OpenCV)
# =========================

def extrair_tabela_super(img: Image.Image) -> tuple[pd.DataFrame | None, list[str]]:
    logs: list[str] = []

    img_cv = np.array(img)

    # FIX crítico: PIL entrega RGB — converter corretamente para escala de cinza
    gray = cv2.cvtColor(img_cv, cv2.COLOR_RGB2GRAY)  # era BGR2GRAY — invertia canais R e B

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
        if abs(y - last_y) > 10:
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
            text = pytesseract.image_to_string(cell, config="--oem 3 --psm 6")
            linha.append(text.strip())
        tabela.append(linha)

    if tabela:
        df = normalizar_colunas(pd.DataFrame(tabela))
        logs.append("Tabela detectada via OpenCV + OCR")
        return df, logs

    return None, logs


# =========================
# OCR — fallback simples (linha por linha)
# =========================

def extrair_tabela_ocr(img: Image.Image) -> tuple[pd.DataFrame | None, list[str]]:
    logs: list[str] = []

    img_cv = np.array(img)

    # FIX crítico: mesmo bug de canal de cor corrigido aqui
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
    """
    Extrai tabelas de PDF ou imagem.
    Cacheado por conteúdo — não reprocessa o mesmo arquivo em reruns.
    """
    tabelas: list[pd.DataFrame] = []
    logs: list[str] = []

    if "pdf" in tipo:
        tb, log_pdf = processar_pdf(bytes_file)
        tabelas.extend(tb)
        logs.extend(log_pdf)

        if not tabelas:
            logs.append("Nenhuma tabela via pdfplumber — iniciando fallback OCR")
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

def gerar_excel(tabelas: list[pd.DataFrame]) -> bytes:
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for i, df in enumerate(tabelas):
            nome = re.sub(r"[\\/*?:\[\]]", "", f"Tabela_{i+1}")[:31]
            # normalizar_colunas já foi chamado no pipeline — não repete aqui
            df.to_excel(writer, index=False, sheet_name=nome)

    return buffer.getvalue()  # FIX: era seek(0) + read() — getvalue() é mais direto e seguro


# =========================
# UI PRINCIPAL
# =========================

st.title("📄 Extrator Inteligente PRO")
st.caption("PDFs e imagens → Excel em lote")

arquivos = st.file_uploader(
    "Envie PDFs ou imagens",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if arquivos:
    st.info(f"{len(arquivos)} arquivo(s) carregado(s)")

    if st.button("🚀 Iniciar processamento"):
        progresso = st.progress(0)
        status = st.empty()
        total = len(arquivos)

        # Reseta resultados da sessão
        st.session_state.resultados = []

        for i, arquivo in enumerate(arquivos):
            status.text(f"Processando {i+1}/{total} → {arquivo.name}")

            if arquivo.size > MAX_SIZE_MB * 1024 * 1024:
                st.error(f"❌ {arquivo.name} excede {MAX_SIZE_MB}MB")
                progresso.progress((i + 1) / total)
                continue

            try:
                tabelas, logs = processar_arquivo(arquivo.getvalue(), arquivo.type)
            except Exception as e:
                st.error(f"❌ Erro em {arquivo.name}: {e}")
                progresso.progress((i + 1) / total)
                continue

            if tabelas:
                excel = gerar_excel(tabelas)

                # Guarda no session_state — os botões de download são renderizados
                # fora do bloco do botão, evitando o sumiço em reruns
                st.session_state.resultados.append({
                    "nome": arquivo.name,
                    "qtd": len(tabelas),
                    "excel": excel,
                    "tabelas": tabelas,
                    "logs": logs,
                })

                # Também salva no histórico (limitado)
                st.session_state.historico.append({
                    "nome": arquivo.name,
                    "data": datetime.now().strftime("%d/%m %H:%M"),
                    "qtd": len(tabelas),
                    "excel": excel,
                })
                st.session_state.historico = st.session_state.historico[-MAX_HISTORICO:]

            else:
                st.warning(f"⚠️ {arquivo.name}: nenhuma tabela encontrada")

            progresso.progress((i + 1) / total)
            # FIX: time.sleep(0.3) removido — não tinha propósito

        status.success("✅ Concluído!")

# =========================
# RESULTADOS — renderizados fora do bloco do botão para não somírem
# =========================

if st.session_state.resultados:
    st.markdown("## 📊 Resultados")

    # Botão ZIP
    if len(st.session_state.resultados) > 1:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for r in st.session_state.resultados:
                nome_limpo = r["nome"].rsplit(".", 1)[0]
                zf.writestr(f"{nome_limpo}.xlsx", r["excel"])
        st.download_button(
            "📦 Baixar TODOS (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="tabelas_extraidas.zip"
        )

    for item in st.session_state.resultados:
        st.markdown(f"**📄 {item['nome']}** — {item['qtd']} tabela(s)")

        st.download_button(
            f"⬇️ Baixar Excel — {item['nome']}",
            data=item["excel"],
            file_name=f"{item['nome'].rsplit('.', 1)[0]}.xlsx",
            key=f"dl_{item['nome']}"
        )

        with st.expander(f"🔍 Preview — {item['nome']}"):
            for df in item["tabelas"]:
                st.dataframe(df, use_container_width=True)

        with st.expander(f"📜 Logs — {item['nome']}"):
            for log in item["logs"]:
                st.text(log)

        st.divider()

# =========================
# HISTÓRICO — session_state em vez de arquivo JSON no disco
# =========================

st.markdown("## 📜 Histórico")

if st.session_state.historico:
    for item in reversed(st.session_state.historico):
        st.markdown(f"**📄 {item['nome']}**")
        st.caption(f"{item['data']} • {item['qtd']} tabelas")

        st.download_button(
            "⬇️ Baixar novamente",
            data=item["excel"],
            file_name=f"{item['nome'].rsplit('.', 1)[0]}.xlsx",
            key=f"hist_{item['nome']}_{item['data']}"
        )

        st.divider()
else:
    st.info("Nenhum histórico ainda.")
