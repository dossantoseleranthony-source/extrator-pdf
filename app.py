import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import json
import os
import time
import base64
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
MAX_HISTORICO = 20

st.set_page_config(page_title="Extrator Inteligente PRO", layout="wide")

# =========================
# HISTÓRICO
# =========================
def salvar_historico(nome, qtd, excel_bytes):
    registro = {
        "arquivo": nome,
        "qtd_tabelas": qtd,
        "data": datetime.now().strftime("%d/%m %H:%M"),
        "excel": base64.b64encode(excel_bytes).decode("utf-8")
    }

    if os.path.exists("historico.json"):
        with open("historico.json", "r") as f:
            dados = json.load(f)
    else:
        dados = []

    dados.append(registro)
    dados = dados[-MAX_HISTORICO:]

    with open("historico.json", "w") as f:
        json.dump(dados, f, indent=4)

def carregar_historico():
    if not os.path.exists("historico.json"):
        return []
    with open("historico.json", "r") as f:
        return json.load(f)

# =========================
# LIMPEZA
# =========================
def limpar_df(df):
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    return df.fillna("")

# 🔥 NOVO: evitar colunas duplicadas
def garantir_colunas_unicas(df):
    cols = []
    for i, col in enumerate(df.columns):
        col = str(col).strip() or f"col_{i}"

        if col in cols:
            col = f"{col}_{i}"

        cols.append(col)

    df.columns = cols
    return df

# =========================
# PDF
# =========================
def processar_pdf(pdf_bytes):
    tabelas, logs = [], []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pagina in enumerate(pdf.pages):
            tb = pagina.extract_tables() or []
            logs.append(f"Página {i+1}: {len(tb)} tabela(s)")

            for tabela in tb:
                df = pd.DataFrame(tabela)

                if df.empty or df.shape[1] <= 1:
                    continue

                df = limpar_df(df)

                if df.shape[0] > 1:
                    header = df.iloc[0].astype(str)
                    if header.str.len().mean() > 2:
                        df.columns = header
                        df = df.iloc[1:]

                df = garantir_colunas_unicas(df)
                tabelas.append(df)

    return tabelas, logs

# =========================
# OCR SUPER
# =========================
def extrair_tabela_super(img):
    logs = []

    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

    thresh = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        15, 4
    )

    horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (40, 1))
    horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel)

    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, 40))
    vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel)

    table_mask = cv2.add(horizontal, vertical)

    contours, _ = cv2.findContours(table_mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)

    boxes = [cv2.boundingRect(c) for c in contours]
    boxes = sorted(boxes, key=lambda b: (b[1], b[0]))

    rows = []
    current_row = []
    last_y = -1

    for (x, y, w, h) in boxes:
        if last_y == -1:
            last_y = y

        if abs(y - last_y) > 10:
            rows.append(current_row)
            current_row = []
            last_y = y

        current_row.append((x, y, w, h))

    if current_row:
        rows.append(current_row)

    tabela = []

    for row in rows:
        row = sorted(row, key=lambda b: b[0])
        linha = []

        for (x, y, w, h) in row:
            cell = img_cv[y:y+h, x:x+w]
            text = pytesseract.image_to_string(cell, config="--oem 3 --psm 6")
            linha.append(text.strip())

        tabela.append(linha)

    if tabela:
        df = pd.DataFrame(tabela)
        df = garantir_colunas_unicas(df)
        logs.append("Tabela detectada com estrutura (OpenCV + OCR)")
        return df, logs

    return None, logs

# =========================
# OCR fallback
# =========================
def extrair_tabela_ocr(img):
    logs = []

    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

    data = pytesseract.image_to_data(
        gray,
        output_type=pytesseract.Output.DATAFRAME
    ).dropna()

    linhas = data.groupby("line_num")

    tabela = []
    max_cols = 0

    for _, linha in linhas:
        palavras = linha.sort_values("left")
        row = palavras["text"].tolist()

        if row:
            tabela.append(row)
            max_cols = max(max_cols, len(row))

    tabela_pad = [r + [""] * (max_cols - len(r)) for r in tabela]

    if tabela_pad:
        df = pd.DataFrame(tabela_pad)
        df = garantir_colunas_unicas(df)
        logs.append("OCR simples aplicado")
        return df, logs

    return None, logs

# =========================
# PIPELINE
# =========================
@st.cache_data(show_spinner=False)
def processar_arquivo(bytes_file, tipo):
    tabelas, logs = [], []

    if "pdf" in tipo:
        tb, log_pdf = processar_pdf(bytes_file)
        tabelas.extend(tb)
        logs.extend(log_pdf)

        if not tabelas:
            logs.append("Fallback OCR com detecção de tabela")

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
def gerar_excel(tabelas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for i, df in enumerate(tabelas):
            nome = f"Tabela_{i+1}"
            nome = re.sub(r"[\\/*?:\[\]]", "", nome)[:31]
            df = garantir_colunas_unicas(df)
            df.to_excel(writer, index=False, sheet_name=nome)

    buffer.seek(0)
    return buffer.read()

# =========================
# UI
# =========================
st.title("📄 Extrator Inteligente")
st.caption("Download em lote")

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
        arquivos_excel = []

        for i, arquivo in enumerate(arquivos):

            status.text(f"Processando {i+1}/{total} → {arquivo.name}")

            if arquivo.size > MAX_SIZE_MB * 1024 * 1024:
                st.error(f"{arquivo.name} excede {MAX_SIZE_MB}MB")
                continue

            try:
                tabelas, logs = processar_arquivo(
                    arquivo.getvalue(),
                    arquivo.type
                )
            except Exception as e:
                st.error(f"Erro em {arquivo.name}: {str(e)}")
                continue

            if tabelas:
                excel = gerar_excel(tabelas)
                arquivos_excel.append((arquivo.name, excel))

                salvar_historico(arquivo.name, len(tabelas), excel)

                st.download_button(
                    f"⬇️ {arquivo.name}",
                    data=excel,
                    file_name=f"{arquivo.name}.xlsx"
                )

                with st.expander(f"📊 Preview - {arquivo.name}"):
                    for df in tabelas:
                        df = garantir_colunas_unicas(df)
                        st.dataframe(df, use_container_width=True)

                with st.expander(f"📜 Logs - {arquivo.name}"):
                    for log in logs:
                        st.text(log)

            progresso.progress((i + 1) / total)
            time.sleep(0.3)

        if arquivos_excel:
            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for nome, excel_bytes in arquivos_excel:
                    nome_limpo = nome.split(".")[0]
                    zf.writestr(f"{nome_limpo}.xlsx", excel_bytes)

            zip_buffer.seek(0)

            st.download_button(
                "📦 Baixar TODOS (ZIP)",
                data=zip_buffer,
                file_name="tabelas_extraidas.zip"
            )

        status.success("✅ Concluído!")

# =========================
# HISTÓRICO
# =========================
st.markdown("## 📜 Histórico")

historico = carregar_historico()

if historico:
    for item in historico[::-1]:
        st.markdown(f"### 📄 {item['arquivo']}")
        st.caption(f"{item['data']} • {item['qtd_tabelas']} tabelas")

        excel_bytes = base64.b64decode(item["excel"])

        st.download_button(
            "⬇️ Baixar novamente",
            data=excel_bytes,
            file_name=f"{item['arquivo']}.xlsx"
        )
else:
    st.info("Nenhum histórico ainda.")
