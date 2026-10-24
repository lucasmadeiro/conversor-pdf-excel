import io
import pandas as pd
import pdfplumber
import streamlit as st

# Configuração da Página com layout wide para aproveitar melhor o espaço
st.set_page_config(
    page_title="Conversor Profissional de PDF",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilo CSS Personalizado para deixar o visual mais elegante e moderno
st.markdown(
    """
    <style>
        .main {
            background-color: #fafbfc;
        }
        .stButton>button {
            border-radius: 8px;
            font-weight: 600;
            background-color: #2563eb;
            color: white;
            padding: 0.5rem 1rem;
            border: none;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }
        .stButton>button:hover {
            background-color: #1d4ed8;
            color: white;
        }
        .upload-text {
            font-size: 1.1rem;
            color: #4b5563;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
  st.image(
      "https://img.icons8.com/clouds/200/microsoft-excel-2019.png", width=100
  )
  st.title("Painel de Controle")
  st.markdown("---")
  st.markdown("### 💡 Dicas de Uso")
  st.info(
      "• **Extratos Bancários:** O sistema alinha o texto preservando a"
      " estrutura original.\n• **Editais e Textos:** O conteúdo é convertido"
      " linha por linha de forma limpa."
  )
  st.markdown("---")
  st.markdown(
      "<p style='text-align: center; color: gray; font-size: 0.8rem;'>Desenvolvido"
      " com Streamlit & Python</p>",
      unsafe_allow_html=True,
  )

# --- CORPO PRINCIPAL ---
st.title("⚡ Conversor Inteligente de PDF para Excel")
st.markdown(
    "<p class='upload-text'>Transforme extratos, relatórios e documentos"
    " complexos em planilhas Excel organizadas instantaneamente.</p>",
    unsafe_allow_html=True,
)

st.markdown("---")

# Container central de Upload com design elegante
col1, col2, col3 = st.columns([1, 6, 1])

with col2:
  uploaded_file = st.file_uploader(
      "Arraste ou selecione o seu ficheiro PDF", type=["pdf"], key="pdf"
  )

  if uploaded_file is not None:
    # Barra de progresso fluida
    progress_bar = st.progress(0)
    status_text = st.empty()

    status_text.text("A ler o documento PDF...")
    progress_bar.progress(30)

    try:
      with pdfplumber.open(uploaded_file) as pdf:
        linhas_extraidas = []

        for num_pag, page in enumerate(pdf.pages):
          linhas_extraidas.append([f"--- PÁGINA {num_pag + 1} ---"])
          texto_pagina = page.extract_text(layout=True)
          if texto_pagina:
            for linha in texto_pagina.split("\n"):
              if linha.strip():
                linhas_extraidas.append([linha.strip()])

        progress_bar.progress(80)
        status_text.text("A formatar os dados para a planilha...")

        if linhas_extraidas:
          df = pd.DataFrame(linhas_extraidas, columns=["Conteúdo do Documento"])

          progress_bar.progress(100)
          status_text.empty()
          progress_bar.empty()

          # Alerta de Sucesso
          st.success(
              f"✨ Conversão concluída com sucesso! Total de linhas processadas:"
              f" {len(df)}"
          )

          # Criação de Abas (Tabs) para uma navegação limpa e moderna
          tab1, tab2 = st.tabs(
              ["📋 Pré-visualização da Planilha", "📥 Central de Download"]
          )

          with tab1:
            st.markdown("### Visualização prévia dos dados extraídos:")
            st.dataframe(df, use_container_width=True, height=400)

          with tab2:
            st.markdown("### O seu ficheiro está pronto para ser baixado!")
            st.write(
                "Clique no botão abaixo para descarregar o documento convertido"
                " no formato oficial do Excel (`.xlsx`)."
            )

            # Prepara o ficheiro Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              df.to_excel(writer, index=False, sheet_name="Extrato_Organizado")
            excel_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=excel_data,
                file_name="extrato_convertido_profissional.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        else:
          progress_bar.empty()
          status_text.empty()
          st.warning(
              "⚠️ Não foi possível extrair texto legível deste documento."
          )

    except Exception as e:
      progress_bar.empty()
      status_text.empty()
      st.error(f"❌ Ocorreu um erro ao processar o ficheiro: {e}")
