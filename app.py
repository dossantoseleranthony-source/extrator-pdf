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

st.markdown("""
<style>
.main { background: linear-gradient(135deg, #0f172a, #1e293b); }
h1, h2, h3 { color: #e2e8f0; }
.stButton>button {
    background: linear-gradient(90deg, #00c8ff, #007cf0);
    color: white; border-radius: 8px; border: none;
    font-weight: bold; transition: 0.3s;
}
.stButton>button:hover { transform: scale(1.05); }
.card {
    background: #111827; padding: 15px; border-radius: 12px;
    box-shadow: 0px 0px 10px rgba(0,200,255,0.1); margin-bottom: 10px;
}
section[data-testid="stSidebar"] { background-color: #020617; }
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

if "historico" not in st.session_state:
    st.session_state.historico = []

# =========================
# FUNÇÕES
# =========================

def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Limpa nomes de colunas e garante unicidade — unifica corrigir_colunas + garantir_colunas_unicas."""
    vistas = {}
    novas = []
    for i, col in enumerate(df.columns):
        nome = str(col).strip().replace("\n", " ")
        if not nome or nome.lower() == "none" or len(nome) > 40:
            nome = f"col_{i}"
        # Desambigua duplicatas com sufixo incremental
        count = vistas.get(nome, 0)
        vistas[nome] = count + 1
        novas.append(f"{nome}_{count}" if count else nome)
    df.columns = novas
    return df


def limpar_df(df: pd.DataFrame) -> pd.DataFrame:
    """Remove linhas/colunas vazias, preenche NaN e filtra linhas de paginação."""
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)  # sem inplace — retorna novo df
    df = df.fillna("")
    mascara_pag = df.apply(lambda row: row.astype(str).str.contains(r"\bpág", case=False).any(), axis=1)
    return df[~mascara_pag].reset_index(drop=True)


@st.cache_data(show_spinner=False)
def processar_pdf(pdf_bytes: bytes) -> tuple[list[pd.DataFrame], list[str]]:
    """Extrai tabelas do PDF. Resultado cacheado por conteúdo do arquivo."""
    todas_tabelas: list[pd.DataFrame] = []
    logs: list[str] = []

    config_linhas = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_y_tolerance": 15,
    }
    config_texto = {"vertical_strategy": "text", "horizontal_strategy": "text"}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for i, pagina in enumerate(pdf.pages):
            tabelas = pagina.extract_tables(config_linhas) or pagina.extract_tables(config_texto)
            logs.append(f"Página {i+1}: {len(tabelas) if tabelas else 0} tabela(s)")

            for tabela in (tabelas or []):
                df = pd.DataFrame(tabela)
                if df.empty:
                    continue

                df = limpar_df(df)
                if df.empty or df.shape[1] <= 1:
                    continue

                # Promove primeira linha a cabeçalho se parecer com títulos
                primeira = df.iloc[0].astype(str)
                if any(len(v) > 2 for v in primeira):
                    df.columns = primeira
                    df = df.iloc[1:].reset_index(drop=True)

                df = normalizar_colunas(df)  # apenas uma chamada, já garante unicidade
                todas_tabelas.append(df)

    return todas_tabelas, logs


@st.cache_data(show_spinner=False)
def gerar_excel(pdf_bytes: bytes, tabelas_json: str) -> bytes:
    """
    Gera Excel a partir das tabelas. Cacheado — evita regerar o mesmo arquivo.
    Recebe tabelas_json como string para ser hashável pelo cache do Streamlit.
    """
    tabelas = [pd.read_json(io.StringIO(t)) for t in pd.read_json(io.StringIO(tabelas_json))]
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for i, df in enumerate(tabelas):
            nome = re.sub(r"[\\/*?:\[\]]", "", f"Tabela_{i+1}")[:31]
            df.to_excel(writer, index=False, sheet_name=nome)
    return buffer.getvalue()


def processar_arquivo(arquivo) -> dict | None:
    """Orquestra leitura, extração e geração de Excel de um único arquivo."""
    try:
        pdf_bytes = arquivo.read()
        tabelas, logs = processar_pdf(pdf_bytes)

        if not tabelas:
            st.warning(f"⚠️ {arquivo.name}: nenhuma tabela encontrada")
            return None

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            for i, df in enumerate(tabelas):
                nome = re.sub(r"[\\/*?:\[\]]", "", f"Tabela_{i+1}")[:31]
                df.to_excel(writer, index=False, sheet_name=nome)
        excel_bytes = buffer.getvalue()

        return {
            "nome": arquivo.name,
            "data": datetime.now().strftime("%d/%m %H:%M"),
            "qtd": len(tabelas),
            "arquivo": excel_bytes,          # só bytes — não guarda DataFrames inteiros
            "tabelas": tabelas,              # mantido apenas para preview (pode remover se não precisar)
            "logs": logs,
        }

    except Exception as e:
        st.error(f"❌ Erro ao processar `{arquivo.name}`: {e}")
        return None


# =========================
# UI PRINCIPAL
# =========================
st.title("📄 Extrator Inteligente de Tabelas")
st.caption("Transforme PDFs em Excel em segundos")

arquivos = st.file_uploader("📂 Envie PDFs", type=["pdf"], accept_multiple_files=True)

if arquivos and st.button("🚀 Processar arquivos"):
    progresso = st.progress(0)
    status = st.empty()

    for i, arquivo in enumerate(arquivos):
        status.markdown(f"🔄 **Processando:** `{arquivo.name}`")
        resultado = processar_arquivo(arquivo)
        if resultado:
            st.session_state.historico.append(resultado)
        progresso.progress((i + 1) / len(arquivos))

    status.success("✅ Processamento concluído!")

# =========================
# HISTÓRICO
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

        _, col_btn = st.columns([3, 1])
        with col_btn:
            st.download_button("⬇️ Excel", data=item["arquivo"], file_name=f"{item['nome']}.xlsx")

        if mostrar_preview and item.get("tabelas"):
            with st.expander("🔍 Preview"):
                for df in item["tabelas"]:
                    st.dataframe(df, use_container_width=True)

        if mostrar_logs and item.get("logs"):
            with st.expander("📜 Logs"):
                for log in item["logs"]:
                    st.text(log)

        st.divider()
else:
    st.info("Nenhum arquivo processado ainda.")
