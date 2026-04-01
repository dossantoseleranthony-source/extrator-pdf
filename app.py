import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime

# OCR
import pytesseract
import cv2
import numpy as np
from PIL import Image
from pdf2image import convert_from_bytes

# =========================
# CONFIG OCR (STREAMLIT CLOUD)
# =========================
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract"

# =========================
# CONFIG APP
# =========================
st.set_page_config(page_title="Extrator de Tabelas", layout="wide")

# =========================
# FUNÇÕES AUXILIARES
# =========================
def limpar_df(df):
    df = df.dropna(how='all', axis=1).dropna(how='all', axis=0)
    df = df.fillna("")
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

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pagina in pdf.pages:
            tb = pagina.extract_tables({
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
            }) or pagina.extract_tables({
                "vertical_strategy": "text",
                "horizontal_strategy": "text",
            })

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

    return tabelas


# =========================
# OCR MELHORADO
# =========================
def extrair_tabela_imagem(img):
    try:
        img_cv = np.array(img)
        gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

        gray = cv2.GaussianBlur(gray, (5, 5), 0)

        thresh = cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            11, 2
        )

        data = pytesseract.image_to_data(
            thresh,
            output_type=pytesseract.Output.DATAFRAME
        )

        data = data.dropna()

        linhas = data.groupby("line_num")["text"].apply(list)

        tabela = [linha for linha in linhas if any(str(x).strip() for x in linha)]

        if not tabela:
            return None

        return pd.DataFrame(tabela)

    except:
        return None


# =========================
# PIPELINE PRINCIPAL
# =========================
@st.cache_data(show_spinner=False)
def processar_arquivo(bytes_file, tipo):
    tabelas = []

    if "pdf" in tipo:
        tabelas = processar_pdf(bytes_file)

        # fallback OCR
        if not tabelas:
            imagens = convert_from_bytes(bytes_file)

            for img in imagens:
                df = extrair_tabela_imagem(img)
                if df is not None:
                    tabelas.append(df)

    else:
        img = Image.open(io.BytesIO(bytes_file))
        df = extrair_tabela_imagem(img)

        if df is not None:
            tabelas.append(df)

    return tabelas


# =========================
# EXCEL
# =========================
def gerar_excel(tabelas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
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
st.title("📄 Extrator de Tabelas (PDF + Imagem)")
st.caption("Funciona com PDF normal, escaneado e imagens")

arquivos = st.file_uploader(
    "Envie arquivos",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if arquivos:
    for arquivo in arquivos:

        # 🚫 Limite de tamanho (evita travar)
        if arquivo.size > 10 * 1024 * 1024:
            st.error(f"{arquivo.name} é muito grande (máx 10MB)")
            continue

        st.markdown(f"### 📄 {arquivo.name}")

        with st.spinner("Processando..."):
            tabelas = processar_arquivo(arquivo.getvalue(), arquivo.type)

        if tabelas:
            excel = gerar_excel(tabelas)

            st.download_button(
                "⬇️ Baixar Excel",
                data=excel,
                file_name=f"{arquivo.name}.xlsx"
            )

            with st.expander("📊 Preview"):
                for df in tabelas:
                    st.dataframe(df, use_container_width=True)

        else:
            st.warning("Nenhuma tabela encontrada")
