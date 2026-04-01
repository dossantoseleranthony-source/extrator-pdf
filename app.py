import streamlit as st
import pdfplumber
import pandas as pd
import io
import re

# =========================
# CONFIG STREAMLIT
# =========================
st.set_page_config(layout="wide")
st.title("📄 Extrator Inteligente de Tabelas de PDFs")
st.write("Upload do PDF → Extração → Excel pronto")

# =========================
# FUNÇÃO: CORRIGIR COLUNAS (ROBUSTA)
# =========================
def corrigir_colunas(df):
    novas_cols = []

    for i, col in enumerate(df.columns):
        nome = str(col).strip()

        # remove quebra de linha
        nome = nome.replace("\n", " ")

        # remove nomes absurdos (títulos grandes)
        if len(nome) > 40:
            nome = f"col_{i}"

        # vazio ou None
        if nome == "" or nome.lower() == "none":
            nome = f"col_{i}"

        # evitar duplicados
        if nome in novas_cols:
            nome = f"{nome}_{i}"

        novas_cols.append(nome)

    df.columns = novas_cols
    return df

# =========================
# GARANTIR COLUNAS ÚNICAS (ANTI-ERRO FINAL)
# =========================
def garantir_colunas_unicas(df):
    cols = []
    for i, col in enumerate(df.columns):
        col = str(col)
        if col in cols:
            col = f"{col}_{i}"
        cols.append(col)
    df.columns = cols
    return df

# =========================
# LIMPEZA DE DADOS
# =========================
def limpar_df(df):
    df.dropna(how='all', axis=1, inplace=True)
    df.dropna(how='all', axis=0, inplace=True)
    df.fillna("", inplace=True)

    # remover rodapé tipo "Página X"
    df = df[~df.apply(lambda row: row.astype(str).str.contains("pág", case=False).any(), axis=1)]

    return df

# =========================
# PROCESSAMENTO PRINCIPAL
# =========================
@st.cache_data(show_spinner=False)
def processar_pdf(pdf_bytes):
    todas_tabelas = []

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for pagina in pdf.pages:

            # Estratégia 1 (linhas)
            config_linhas = {
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_y_tolerance": 15,
            }

            tabelas = pagina.extract_tables(table_settings=config_linhas)

            # Estratégia 2 (texto)
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

                # tentar detectar header real
                if df.shape[0] > 1:
                    primeira_linha = df.iloc[0].astype(str)

                    if any(len(str(x)) > 2 for x in primeira_linha):
                        df.columns = primeira_linha
                        df = df[1:]

                df = corrigir_colunas(df)

                # ignorar tabelas ruins
                if df.shape[1] <= 1:
                    continue

                todas_tabelas.append(df)

    return todas_tabelas

# =========================
# UPLOAD
# =========================
arquivo_pdf = st.file_uploader("Selecione um PDF", type=["pdf"])

if arquivo_pdf is not None:
    espaco_topo = st.empty()

    try:
        with st.spinner("🔍 Processando PDF..."):
            pdf_bytes = arquivo_pdf.read()
            tabelas = processar_pdf(pdf_bytes)

        if tabelas:
            st.success(f"{len(tabelas)} tabela(s) encontrada(s)")

            for i, df in enumerate(tabelas):
                st.subheader(f"Tabela {i+1}")
                st.dataframe(df, use_container_width=True)

            # =========================
            # EXPORTAR EXCEL
            # =========================
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

            espaco_topo.download_button(
                "📥 Baixar Excel",
                data=buffer.getvalue(),
                file_name="tabelas_extraidas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        else:
            espaco_topo.warning("Nenhuma tabela encontrada.")

    except Exception as e:
        st.error(f"Erro: {e}")