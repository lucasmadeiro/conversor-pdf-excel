import io
import pandas as pd
import pdfplumber
import streamlit as st

# Configuração da Página
st.set_page_config(
    page_title="Conversor Profissional de PDF",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilo CSS Personalizado (incluindo a remoção do texto de limite de tamanho)
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
            width: 100%;
        }
        .stButton>button:hover {
            background-color: #1d4ed8;
            color: white;
        }
        
        /* Personalização do Botão de Upload para Azul Médio */
        [data-testid="stFileUploader"] button {
            background-color: #2563eb !important;
            color: white !important;
            border-radius: 8px !important;
            border: none !important;
        }
        [data-testid="stFileUploader"] button:hover {
            background-color: #1d4ed8 !important;
            color: white !important;
        }

        /* Oculta o texto de limite de tamanho (200MB per file • PDF) */
        [data-testid="stFileUploader"] small {
            display: none !important;
        }

        /* Título personalizado e centralizado para a seleção */
        .upload-label {
            text-align: center;
            color: #2563eb;
            font-size: 1.2rem;
            font-weight: 600;
            margin-bottom: 0.5rem;
        }

        .upload-text {
            font-size: 1.1rem;
            color: #4b5563;
            text-align: center;
        }
    </style>
""",
    unsafe_allow_html=True,
)

# --- BARRA LATERAL ---
with st.sidebar:
  st.image(
      "https://img.icons8.com/clouds/200/microsoft-excel-2019.png", width=100
  )
  st.title("Painel de Controle")
  st.markdown("---")
  st.markdown("### 💡 Como Funciona")
  st.info(
      "1. Selecione o seu ficheiro PDF.\n2. Escolha como apresentar os dados.\n3."
      " Clique em **Iniciar Processamento**."
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
    "<p class='upload-text'>Transforme documentos complexos em planilhas"
    " Excel organizadas do seu jeito.</p>",
    unsafe_allow_html=True,
)

st.markdown("---")

col1, col2, col3 = st.columns([1, 6, 1])

with col2:
  # Rótulo HTML centralizado e destacado em azul logo acima do uploader
  st.markdown(
      "<div class='upload-label'>📁 Selecionar um arquivo PDF</div>",
      unsafe_allow_html=True,
  )

  uploaded_file = st.file_uploader("", type=["pdf"], key="pdf")

  if uploaded_file is not None:
    st.markdown("---")
    st.markdown("### Como apresentar os dados?")

    # Escolha do formato ANTES de processar
    modo_export = st.radio(
        "Selecione o formato de saída:",
        [
            "📑 Tabela única (Todos os dados numa só página contínua)",
            "📂 Várias tabelas (Separar cada página em abas diferentes)",
        ],
        label_visibility="collapsed",
    )

    st.markdown("")

    # Botão explícito para iniciar
    if st.button("🚀 Iniciar Processamento"):
      progress_bar = st.progress(0)
      status_text = st.empty()

      status_text.text("A ler o documento PDF...")
      progress_bar.progress(30)

      try:
        with pdfplumber.open(uploaded_file) as pdf:
          paginas_dict = {}
          todas_linhas = []

          for num_pag, page in enumerate(pdf.pages):
            cabecalho_pag = [f"--- PÁGINA {num_pag + 1} ---"]
            todas_linhas.append(cabecalho_pag)

            linhas_pagina = [f"--- PÁGINA {num_pag + 1} ---"]
            texto_pagina = page.extract_text(layout=True)

            if texto_pagina:
              for linha in texto_pagina.split("\n"):
                if linha.strip():
                  todas_linhas.append([linha.strip()])
                  linhas_pagina.append(linha.strip())

            paginas_dict[f"Pagina_{num_pag + 1}"] = pd.DataFrame(
                linhas_pagina, columns=["Conteúdo do Documento"]
            )

          progress_bar.progress(80)
          status_text.text("A gerar a planilha final...")

          if todas_linhas:
            df_completo = pd.DataFrame(
                todas_linhas, columns=["Conteúdo do Documento"]
            )

            progress_bar.progress(100)
            status_text.empty()
            progress_bar.empty()

            st.success("✨ Processamento concluído com sucesso!")

            # Abas visuais do resultado
            tab1, tab2 = st.tabs(
                ["📋 Pré-visualização", "📥 Central de Download"]
            )

            with tab1:
              st.markdown("### Visualização prévia dos dados:")
              st.dataframe(df_completo, use_container_width=True, height=400)

            with tab2:
              st.markdown("### O seu ficheiro está pronto para ser baixado!")

              output = io.BytesIO()
              with pd.ExcelWriter(output, engine="openpyxl") as writer:
                if "Tabela única" in modo_export:
                  df_completo.to_excel(
                      writer, index=False, sheet_name="Documento_Completo"
                  )
                else:
                  for nome_aba, df_p in paginas_dict.items():
                    df_p.to_excel(writer, index=False, sheet_name=nome_aba)

              excel_data = output.getvalue()

              st.download_button(
                  label="📥 Baixar Planilha Excel (.xlsx)",
                  data=excel_data,
                  file_name="documento_convertido.xlsx",
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
