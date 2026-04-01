import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime

# OCR / IMAGEM
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_bytes

# =========================
# CONFIG
# =========================
st.set_page_config(page_title="Extrator Inteligente", layout="wide")

# =========================
# FUNÇÕES BASE
# =========================
def limpar_df(df):
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    df = df.fillna("")

    mask = df.astype(str).apply(lambda col: col.str.contains("pág", case=False, na=False))
    df = df[~mask.any(axis=1)]

    return df


def corrigir_colunas(df):
    novas = []
    for i, col in enumerate(df.columns):
        nome = str(col).strip()

        if not nome or len(nome) > 40:
            nome = f"col_{i}"

        if nome in novas:
            nome = f"{nome}_{i}"

        novas.append(nome)

    df.columns = novas
    return df


def garantir_colunas_unicas(df):
    cols = []
    for i, col in enumerate(df.columns):
        if col in cols:
            col = f"{col}_{i}"
        cols.append(col)
    df.columns = cols
    return df


# =========================
# PDF NORMAL
# =========================
def processar_pdf(pdf_bytes):
    tabelas = []
    logs = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pagina in enumerate(pdf.pages):

            tb = pagina.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
            }) or pagina.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
            })

            logs.append(f"Página {i+1}: {len(tb) if tb else 0} tabela(s)")

            if not tb:
                continue

            for tabela in tb:
                df = pd.DataFrame(tabela)

                if df.empty or df.shape[1] <= 1:
                    continue

                df = limpar_df(df)
                df = corrigir_colunas(df)

                if df.shape[0] > 1:
                    primeira = df.iloc[0].astype(str)
                    if primeira.str.len().mean() > 2:
                        df.columns = primeira
                        df = df.iloc[1:]

                tabelas.append(corrigir_colunas(df))

    return tabelas, logs


# =========================
# OCR (IMAGEM)
# =========================
def extrair_tabela_imagem(img):
    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

    texto = pytesseract.image_to_string(thresh, lang="por")

    linhas = texto.split("\n")
    dados = [linha.split() for linha in linhas if linha.strip()]

    if not dados:
        return None

    return pd.DataFrame(dados)


def pdf_para_imagens(pdf_bytes):
    return convert_from_bytes(pdf_bytes)


# =========================
# PIPELINE INTELIGENTE
# =========================
@st.cache_data(show_spinner=False)
def processar_arquivo(pdf_bytes):
    tabelas, logs = processar_pdf(pdf_bytes)

    # fallback OCR
    if not tabelas:
        imagens = pdf_para_imagens(pdf_bytes)

        for i, img in enumerate(imagens):
            df = extrair_tabela_imagem(img)

            if df is not None and not df.empty:
                tabelas.append(df)
                logs.append(f"Página {i+1}: tabela via OCR")

    return tabelas, logs


# =========================
# EXCEL
# =========================
def gerar_excel(tabelas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for i, df in enumerate(tabelas):
            df = garantir_colunas_unicas(df)

            nome = f"Tabela_{i+1}"
            nome = re.sub(r"[\\/*?:\[\]]", "", nome)[:31]

            df.to_excel(writer, index=False, sheet_name=nome)

    buffer.seek(0)
    return buffer.read()


# =========================
# UI
# =========================
st.title("📄 Extrator Inteligente de Tabelas")
st.caption("Agora com suporte a PDFs escaneados e imagens 🚀")

arquivos = st.file_uploader(
    "Envie PDFs ou imagens",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if "historico" not in st.session_state:
    st.session_state.historico = []

if arquivos:
    if st.button("🚀 Processar"):

        for arquivo in arquivos:
            bytes_file = arquivo.getvalue()

            # Se for imagem
            if arquivo.type.startswith("image"):
                img = Image.open(io.BytesIO(bytes_file))
                df = extrair_tabela_imagem(img)

                tabelas = [df] if df is not None else []
                logs = ["Imagem processada via OCR"]

            else:
                tabelas, logs = processar_arquivo(bytes_file)

            if tabelas:
                excel = gerar_excel(tabelas)

                st.session_state.historico.append({
                    "nome": arquivo.name,
                    "tabelas": tabelas,
                    "logs": logs,
                    "arquivo": excel,
                    "data": datetime.now().strftime("%d/%m %H:%M")
                })
            else:
                st.warning(f"{arquivo.name}: nenhuma tabela encontrada")


# =========================
# HISTÓRICO
# =========================
st.markdown("## 📜 Histórico")

for item in reversed(st.session_state.historico):
    st.markdown(f"### 📄 {item['nome']} ({item['data']})")

    st.download_button(
        "⬇️ Baixar Excel",
        data=item["arquivo"],
        file_name=f"{item['nome']}.xlsx"
    )

    with st.expander("📊 Preview"):
        for df in item["tabelas"]:
            st.dataframe(df, use_container_width=True)

    with st.expander("📜 Logs"):
        for log in item["logs"]:
            st.text(log)
