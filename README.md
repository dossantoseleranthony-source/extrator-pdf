# Extrator PDF 📄

Uma aplicação web desenvolvida em **Python** e **Streamlit** projetada para extrair, processar e analisar dados de arquivos PDF de forma prática, rápida e automatizada.

---

## 🚀 O que o aplicativo faz?
O **Extrator PDF** serve como uma ferramenta centralizada para o tratamento de documentos em formato PDF. Ele automatiza o fluxo de leitura de arquivos, permitindo extrair textos e conteúdos que muitas vezes exigem processos manuais complexos, facilitando o dia a dia em tarefas de auditoria, leitura de relatórios ou extração de dados estruturados.

---

## ⚙️ Como ele funciona?
A aplicação combina uma interface web amigável com poderosas bibliotecas e utilitários de sistema para o processamento de documentos:

1. **Interface Web Interativa (`Streamlit`):** Permite que o usuário faça o upload de arquivos PDF diretamente pelo navegador de forma simples e intuitiva, com suporte configurado para arquivos maiores.
2. **Processamento Avançado e OCR (`Tesseract & Poppler`):** Nos bastidores, a aplicação utiliza ferramentas robustas de manipulação de PDF e OCR (Reconhecimento Óptico de Caracteres) através do *Poppler-utils* e *Tesseract-ocr*, garantindo a leitura tanto de PDFs nativos quanto de documentos digitalizados (imagens).
3. **Ambiente Padronizado (`Devcontainer`):** O projeto conta com suporte a *Devcontainers*, permitindo configurar o ambiente de desenvolvimento completo de forma isolada com todas as dependências do sistema já prontas para uso.

---

## 🛠️ Tecnologias Utilizadas

* **[Python](https://www.python.org/)** — Linguagem principal de programação
* **[Streamlit](https://streamlit.io/)** — Framework para construção da interface web
* **[Tesseract OCR](https://github.com/tesseract-ocr/tesseract)** — Motor de reconhecimento óptico de caracteres
* **[Poppler Utils](https://poppler.freedesktop.org/)** — Utilitários para renderização e manipulação de arquivos PDF

---

## 📦 Instalação e Execução Local

Se você deseja rodar o projeto localmente, siga os passos abaixo:

1. **Clone o repositório:**
   ```bash
   git clone https://github.com/SEU-USUARIO/SEU-REPOSITORIO.git
   cd SEU-REPOSITORIO
   ```

2. **Instale as dependências do sistema (necessárias para OCR e manipulação de PDF):**
   * *No Ubuntu/Debian:*
     ```bash
     sudo apt-get update && sudo apt-get install -y tesseract-ocr poppler-utils
     ```

3. **Instale as dependências do Python:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Inicie a aplicação Streamlit:**
   ```bash
   streamlit run app.py
   ```

---

## 👤 Autor

Desenvolvido por **Anthony**.
