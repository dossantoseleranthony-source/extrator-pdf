import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from datetime import datetime

# =========================
# CONFIG UI
# =========================
st.set_page_config(page_title="Extrator PDF", page_icon="📄", layout="wide")

# =========================
# SIDEBAR (CONFIGURAÇÕES)
# =========================
with st.sidebar:
    st.header("⚙️ Configurações")

    mostrar_preview = st.checkbox("Mostrar preview das tabelas", True)
    mostrar_logs = st.checkbox("Mostrar logs de extração", True)

    st.divider()

    if st.button("🗑️ Limpar histórico"):
        st.session_state.historico = []
        st.success("Histórico limpo!")

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


def processar_pdf(pdf_bytes):
    todas_tabelas = []
    logs = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pagina in enumerate(pdf.pages):
            pagina_log = f"Página {i+1}: "

            config_linhas = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 15,
            }

            tabelas = pagina.extract_tables(table_settings=config_linhas)

            if tabelas:
                pagina_log += f"{len(tabelas)} tabela(s) detectada(s) (modo linhas)"
            else:
                config_texto = {
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "intersection_tolerance": 10,
                }
                tabelas = pagina.extract_tables(table_settings=config_texto)

                if tabelas:
                    pagina_log += f"{len(tabelas)} tabela(s) detectada(s) (modo texto)"
                else:
                    pagina_log += "nenhuma tabela encontrada"

            logs.append(pagina_log)

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

    return todas_tabelas, logs


def gerar_excel(tabelas):
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for i, df in enumerate(tabelas):
            df = garantir_colunas_unicas(df)
            nome = f"Tabela_{i+1}"
            nome = re.sub(r"[\\/*?:\[\]]", "", nome)[:31]
            df.to_excel(writer, index=False, sheet_name=nome)

    return buffer.getvalue()

# =========================
# UI PRINCIPAL
# =========================
st.title("📄 Extrator Inteligente de Tabelas")
st.caption("Upload → Análise → Download estruturado")

arquivos = st.file_uploader(
    "📂 Envie PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if arquivos:
    if st.button("🚀 Processar arquivos"):
        progresso = st.progress(0)
        status_text = st.empty()

        for i, arquivo in enumerate(arquivos):
            status_text.text(f"Processando: {arquivo.name}")

            pdf_bytes = arquivo.read()
            tabelas, logs = processar_pdf(pdf_bytes)

            if not tabelas:
                st.warning(f"⚠️ {arquivo.name}: nenhuma tabela encontrada")

            else:
                excel_bytes = gerar_excel(tabelas)

                # salvar histórico
                st.session_state.historico.append({
                    "nome": arquivo.name,
                    "data": datetime.now().strftime("%d/%m %H:%M"),
                    "qtd": len(tabelas),
                    "arquivo": excel_bytes,
                    "tabelas": tabelas,
                    "logs": logs
                })

            progresso.progress((i + 1) / len(arquivos))

        status_text.text("✅ Processamento concluído!")

# =========================
# HISTÓRICO COM PREVIEW + LOG
# =========================
st.markdown("## 📜 Histórico")

if st.session_state.historico:
    for item in reversed(st.session_state.historico):

        st.markdown(f"### 📄 {item['nome']}")
        st.caption(f"🕒 {item['data']} • 📊 {item['qtd']} tabelas")

        col1, col2 = st.columns([3, 1])

        with col2:
            st.download_button(
                "⬇️ Baixar Excel",
                data=item["arquivo"],
                file_name=f"{item['nome']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # 🔍 PREVIEW
        if mostrar_preview:
            with st.expander("🔍 Preview das tabelas"):
                for i, df in enumerate(item["tabelas"]):
                    st.markdown(f"**Tabela {i+1}**")
                    st.dataframe(df, use_container_width=True)

        # 📜 LOG
        if mostrar_logs:
            with st.expander("📜 Log de extração"):
                for log in item["logs"]:
                    st.text(log)

        st.divider()

else:
    st.info("Nenhum arquivo processado ainda.")
