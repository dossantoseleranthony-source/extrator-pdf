import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime

# =========================
# CONFIG UI
# =========================
st.set_page_config(
    page_title="Extrator de PDFs",
    page_icon="📄",
    layout="wide"
)

st.markdown("""
    <style>
        .main {
            background-color: #0e1117;
        }
        .stButton>button {
            background-color: #00c8ff;
            color: black;
            border-radius: 10px;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📄 Extrator Inteligente de Tabelas")
st.caption("Upload → Extração → Excel em segundos")

# =========================
# HISTÓRICO
# =========================
if "historico" not in st.session_state:
    st.session_state.historico = []

# =========================
# FUNÇÕES
# =========================
def corrigir_colunas(df):
    novas_cols = []
    for i, col in enumerate(df.columns):
        nome = str(col).strip().replace("\n", " ")

        if len(nome) > 40 or nome == "" or nome.lower() == "none":
            nome = f"col_{i}"

        if nome in novas_cols:
            nome = f"{nome}_{i}"

        novas_cols.append(nome)

    df.columns = novas_cols
    return df


def garantir_colunas_unicas(df):
    cols = []
    for i, col in enumerate(df.columns):
        col = str(col)
        if col in cols:
            col = f"{col}_{i}"
        cols.append(col)
    df.columns = cols
    return df


def limpar_df(df):
    df.dropna(how='all', axis=1, inplace=True)
    df.dropna(how='all', axis=0, inplace=True)
    df.fillna("", inplace=True)

    df = df[~df.apply(lambda row: row.astype(str).str.contains("pág", case=False).any(), axis=1)]
    return df


@st.cache_data(show_spinner=False)
def processar_pdf(pdf_bytes):
    todas_tabelas = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pagina in pdf.pages:

            config_linhas = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 15,
            }

            tabelas = pagina.extract_tables(table_settings=config_linhas)

            if not tabelas:
                config_texto = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_tolerance": 10,
                }
                tabelas = pagina.extract_tables(table_settings=config_texto)

            for tabela in tabelas:
                df = pd.DataFrame(tabela)

                if df.empty:
                    continue

                df = limpar_df(df)
                df = corrigir_colunas(df)

                if df.shape[0] > 1:
                    primeira_linha = df.iloc[0].astype(str)

                    if any(len(str(x)) > 2 for x in primeira_linha):
                        df.columns = primeira_linha
                        df = df[1:]

                df = corrigir_colunas(df)

                if df.shape[1] <= 1:
                    continue

                todas_tabelas.append(df)

    return todas_tabelas


def gerar_excel(tabelas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        nomes_usados = set()

        for i, df in enumerate(tabelas):
            df = garantir_colunas_unicas(df)

            nome = f"Tabela_{i+1}"
            nome = re.sub(r"[\\/*?:\[\]]", "", nome)[:31]

            if nome in nomes_usados:
                nome = f"{nome}_{i}"

            nomes_usados.add(nome)

            df.to_excel(writer, index=False, sheet_name=nome)

    return buffer.getvalue()


# =========================
# UPLOAD MULTIPLO
# =========================
arquivos = st.file_uploader(
    "📂 Envie um ou mais PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if arquivos:
    col1, col2 = st.columns([3, 1])

    with col2:
        if st.button("🚀 Processar PDFs"):
            progresso = st.progress(0)

            for i, arquivo in enumerate(arquivos):
                pdf_bytes = arquivo.read()
                tabelas = processar_pdf(pdf_bytes)

                if tabelas:
                    excel_bytes = gerar_excel(tabelas)

                    st.session_state.historico.append({
                        "nome": arquivo.name,
                        "data": datetime.now().strftime("%d/%m %H:%M"),
                        "arquivo": excel_bytes,
                        "qtd": len(tabelas)
                    })

                progresso.progress((i + 1) / len(arquivos))

            st.success("Processamento concluído!")

# =========================
# HISTÓRICO VISUAL
# =========================
st.markdown("## 📜 Histórico")

if st.session_state.historico:
    for item in reversed(st.session_state.historico):
        with st.container():
            col1, col2, col3 = st.columns([4, 2, 2])

            col1.markdown(f"**📄 {item['nome']}**  \n🕒 {item['data']}")
            col2.markdown(f"📊 {item['qtd']} tabelas")

            col3.download_button(
                "⬇️ Baixar Excel",
                data=item["arquivo"],
                file_name=f"{item['nome']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            st.divider()
else:
    st.info("Nenhum arquivo processado ainda.")
