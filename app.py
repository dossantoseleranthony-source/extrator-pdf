import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
import json
import os
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

st.set_page_config(page_title="Extrator Inteligente", layout="wide")

# =========================
# HISTÓRICO
# =========================
def salvar_historico(nome, qtd):
    registro = {
        "arquivo": nome,
        "qtd_tabelas": qtd,
        "data": datetime.now().strftime("%d/%m %H:%M")
    }

    if os.path.exists("historico.json"):
        with open("historico.json", "r") as f:
            dados = json.load(f)
    else:
        dados = []

    dados.append(registro)

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

# =========================
# PDF
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

            logs.append(f"Página {i+1}: {len(tb) if tb else 0} tabela(s) encontradas")

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
# OCR INTELIGENTE (ORGANIZA COLUNAS)
# =========================
def extrair_tabela_estruturada(img):
    logs = []

    img_cv = np.array(img)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)

    data = pytesseract.image_to_data(
        gray,
        output_type=pytesseract.Output.DATAFRAME
    )

    data = data.dropna()

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
        logs.append("Tabela estruturada via OCR (colunas organizadas)")
        return pd.DataFrame(tabela_pad), logs

    return None, logs

# =========================
# PIPELINE
# =========================
@st.cache_data(show_spinner=False)
def processar_arquivo(bytes_file, tipo):
    tabelas = []
    logs = []

    if "pdf" in tipo:
        tb, log_pdf = processar_pdf(bytes_file)
        tabelas.extend(tb)
        logs.extend(log_pdf)

        if not tabelas:
            imagens = convert_from_bytes(bytes_file)
            logs.append("Nenhuma tabela detectada → usando OCR")

            for i, img in enumerate(imagens):
                df, log_ocr = extrair_tabela_estruturada(img)
                logs.extend([f"Página {i+1}: {l}" for l in log_ocr])

                if df is not None:
                    tabelas.append(df)

    else:
        img = Image.open(io.BytesIO(bytes_file))
        df, log_ocr = extrair_tabela_estruturada(img)

        logs.extend(log_ocr)

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
            df.to_excel(writer, index=False, sheet_name=nome)

    buffer.seek(0)
    return buffer.read()

# =========================
# UI
# =========================
st.title("📄 Extrator Inteligente de Tabelas")
st.caption("PDF + Imagem + OCR com organização automática")

arquivos = st.file_uploader(
    "Envie arquivos",
    type=["pdf", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

if arquivos:
    for arquivo in arquivos:

        if arquivo.size > 10 * 1024 * 1024:
            st.error(f"{arquivo.name} muito grande (máx 10MB)")
            continue

        st.markdown(f"## 📄 {arquivo.name}")

        with st.spinner("Processando..."):
            tabelas, logs = processar_arquivo(arquivo.getvalue(), arquivo.type)

        if tabelas:
            salvar_historico(arquivo.name, len(tabelas))

            excel = gerar_excel(tabelas)

            st.download_button(
                "⬇️ Baixar Excel",
                data=excel,
                file_name=f"{arquivo.name}.xlsx"
            )

            with st.expander("📊 Preview das Tabelas"):
                for i, df in enumerate(tabelas):
                    st.write(f"Tabela {i+1}")
                    st.dataframe(df, use_container_width=True)

            with st.expander("📜 Logs de processamento"):
                for log in logs:
                    st.text(log)

        else:
            st.warning("Nenhuma tabela encontrada")

# =========================
# HISTÓRICO
# =========================
st.markdown("## 📜 Histórico")

historico = carregar_historico()

if historico:
    for item in historico[::-1]:
        st.write(f"{item['data']} - {item['arquivo']} ({item['qtd_tabelas']} tabelas)")
else:
    st.info("Nenhum histórico ainda.")
