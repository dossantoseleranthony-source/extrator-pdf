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
# 🎨 ESTILO MODERNO
# =========================
st.markdown("""
<style>

/* Fundo */
.main {
    background: linear-gradient(135deg, #0f172a, #1e293b);
}

/* Títulos */
h1, h2, h3 {
    color: #e2e8f0;
}

/* Botões */
.stButton>button {
    background: linear-gradient(90deg, #00c8ff, #007cf0);
    color: white;
    border-radius: 8px;
    border: none;
    font-weight: bold;
    transition: 0.3s;
}

.stButton>button:hover {
    transform: scale(1.05);
}

/* Cards */
.card {
    background: #111827;
    padding: 15px;
    border-radius: 12px;
    box-shadow: 0px 0px 10px rgba(0,200,255,0.1);
    margin-bottom: 10px;
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: #020617;
}

</style>
""", unsafe_allow_html=True)

# =========================
# SIDEBAR
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
# FUNÇÕES (SEM ALTERAÇÃO)
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
                pagina_log += f"{len(tabelas)} tabela(s)"
            else:
                tabelas = pagina.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                })
                pagina_log += f"{len(tabelas) if tabelas else 0} tabela(s)"

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
st.caption("Transforme PDFs em Excel em segundos")

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
            status_text.markdown(f"🔄 **Processando:** `{arquivo.name}`")

            pdf_bytes = arquivo.read()
            tabelas, logs = processar_pdf(pdf_bytes)

            if not tabelas:
                st.warning(f"⚠️ {arquivo.name}: nenhuma tabela encontrada")

            else:
                excel_bytes = gerar_excel(tabelas)

                st.session_state.historico.append({
                    "nome": arquivo.name,
                    "data": datetime.now().strftime("%d/%m %H:%M"),
                    "qtd": len(tabelas),
                    "arquivo": excel_bytes,
                    "tabelas": tabelas,
                    "logs": logs
                })

            progresso.progress((i + 1) / len(arquivos))

        status_text.success("✅ Processamento concluído!")

# =========================
# HISTÓRICO BONITO
# =========================
st.markdown("## 📜 Histórico")

if st.session_state.historico:
    for item in reversed(st.session_state.historico):

        st.markdown(f"""
        <div class="card">
            <b>📄 {item['nome']}</b><br>
            🕒 {item['data']} • 📊 {item['qtd']} tabelas
        </div>
        """, unsafe_allow_html=True)

        col1, col2 = st.columns([3, 1])

        with col2:
            st.download_button(
                "⬇️ Excel",
                data=item["arquivo"],
                file_name=f"{item['nome']}.xlsx"
            )

        if mostrar_preview:
            with st.expander("🔍 Preview"):
                for i, df in enumerate(item["tabelas"]):
                    st.dataframe(df, use_container_width=True)

        if mostrar_logs:
            with st.expander("📜 Logs"):
                for log in item["logs"]:
                    st.text(log)

        st.divider()

else:
    st.info("Nenhum arquivo processado ainda.")
