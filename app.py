import io
import re
from pdf2image import convert_from_bytes
import pandas as pd
import pytesseract
import pdfplumber
import streamlit as st

# Configuração da Página
st.set_page_config(
    page_title="Conversor Profissional de PDF para Excel",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Estilo CSS Personalizado
st.markdown(
    """
    <style>
        .main .block-container {
            padding-top: 2rem !important;
            padding-bottom: 2rem !important;
        }
        [data-testid="stSidebar"] {
            background-color: #e0f2fe !important;
        }
        [data-testid="stSidebar"] h1, 
        [data-testid="stSidebar"] h2, 
        [data-testid="stSidebar"] h3, 
        [data-testid="stSidebar"] p, 
        [data-testid="stSidebar"] span, 
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] div {
            color: #0f172a !important;
        }
        [data-testid="stSidebar"] .stInfo {
            background-color: #bae6fd !important;
            color: #0369a1 !important;
            border: none;
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
        [data-testid="stFileUploader"] small {
            display: none !important;
        }
        [data-testid="stFileUploader"] section {
            padding: 20px 15px !important;
            border: 2px dashed #93c5fd !important;
            border-radius: 10px !important;
            background-color: #ffffff !important;
        }
        .converter-header {
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 15px;
            margin-bottom: 10px;
        }
        .badge-file {
            background-color: #ffffff;
            padding: 6px 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            font-weight: 700;
            font-size: 1rem;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .badge-pdf { color: #dc2626; border: 2px solid #fee2e2; }
        .badge-excel { color: #16a34a; border: 2px solid #dcfce7; }
        .arrow-icon {
            font-size: 1.2rem;
            color: #2563eb;
            font-weight: bold;
        }
        .upload-text {
            font-size: 1rem;
            color: #4b5563;
            text-align: center;
            margin-bottom: 1rem;
        }
    </style>
""",
    unsafe_allow_html=True,
)


def estruturar_extrato_banco_brasil(texto_linhas):
  """Extrai e separa com precisão os lançamentos do extrato bancário."""
  lancamentos = []
  padrao_data = re.compile(r"^\d{2}/\d{2}/\d{4}")

  for linha in texto_linhas:
    linha = linha.strip()
    if not linha:
      continue

    # Verifica se a linha começa com uma data (indicando uma transação)
    if padrao_data.match(linha):
      partes = linha.split()
      data = partes[0]
      # O resto da linha contém os detalhes
      detalhes = " ".join(partes[1:])
      lancamentos.append({"Data": data, "Descrição / Histórico": detalhes})
    else:
      # Linhas de continuação de histórico (ex: nome do favorecido ou detalhes)
      if lancamentos and not padrao_data.match(linha):
        # Ignora cabeçalhos genéricos
        if (
            "Extrato" not in linha
            and "Agência" not in linha
            and "Cliente" not in linha
        ):
          lancamentos[-1]["Descrição / Histórico"] += " " + linha

  if lancamentos:
    return pd.DataFrame(lancamentos)
  else:
    return pd.DataFrame({"Conteúdo Bruto": texto_linhas})


# --- BARRA LATERAL ---
with st.sidebar:
  st.image("logo.png", use_container_width=True)
  st.markdown("---")
  st.title("Painel de Controle")
  st.markdown("---")
  st.markdown("### 💡 Como Funciona")
  st.info(
      "1. Selecione o ficheiro PDF.\n2. Ative o **OCR** para documentos"
      " escaneados.\n3. Clique em **Iniciar Processamento**."
  )
  st.markdown("---")
  st.markdown(
      "<p style='text-align: center; font-size: 0.8rem;'>Desenvolvido com"
      " Streamlit & Python</p>",
      unsafe_allow_html=True,
  )

# --- CORPO PRINCIPAL ---
st.title("⚡ Conversor Inteligente de PDF para Excel")
st.markdown(
    "<p class='upload-text'>Transforme documentos complexos em planilhas"
    " Excel organizadas do seu jeito.</p>",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="converter-header">
        <div class="badge-file badge-pdf">📄 Arquivo PDF</div>
        <div class="arrow-icon">➔ ⚡ ➔</div>
        <div class="badge-file badge-excel">📊 Planilha Excel</div>
    </div>
""",
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([1, 6, 1])

with col2:
  uploaded_file = st.file_uploader("", type=["pdf"], key="pdf")

  st.markdown("### Configurações de Extração")

  usar_ocr = st.checkbox(
      "🔍 Ativar OCR (Recomendado para PDFs escaneados ou sem texto"
      " selecionável)",
      value=True,
  )

  modo_export = st.radio(
      "Selecione o formato de saída:",
      [
          "📑 Tabela única (Todos os dados numa só página contínua)",
          "📂 Várias tabelas (Separar cada página em abas diferentes)",
      ],
      label_visibility="collapsed",
  )

  st.markdown("")

  if uploaded_file is None:
    st.markdown(
        """
        <div style='background-color: #94a3b8; color: white; text-align: center; padding: 0.6rem 1rem; border-radius: 8px; font-weight: 600; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);'>
            ⏳ Aguardando arquivo
        </div>
        """,
        unsafe_allow_html=True,
    )
  else:
    if st.button("🚀 Iniciar Processamento"):
      progress_bar = st.progress(0)
      status_text = st.empty()

      paginas_dict = {}
      todas_linhas_brutas = []
      file_bytes = uploaded_file.getvalue()

      try:
        if usar_ocr:
          status_text.text("A processar páginas via OCR...")
          progress_bar.progress(30)

          images = convert_from_bytes(file_bytes)
          for num_pag, image in enumerate(images):
            texto_ocr = pytesseract.image_to_string(image, lang="por")
            linhas_pag = []
            if texto_ocr:
              for linha in texto_ocr.split("\n"):
                if linha.strip():
                  todas_linhas_brutas.append(linha.strip())
                  linhas_pag.append(linha.strip())

            paginas_dict[f"Pagina_{num_pag + 1}"] = (
                estruturar_extrato_banco_brasil(linhas_pag)
            )
        else:
          status_text.text("A ler o documento PDF...")
          progress_bar.progress(30)

          with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for num_pag, page in enumerate(pdf.pages):
              texto_pagina = page.extract_text()
              linhas_pag = []
              if texto_pagina:
                for linha in texto_pagina.split("\n"):
                  if linha.strip():
                    todas_linhas_brutas.append(linha.strip())
                    linhas_pag.append(linha.strip())

              paginas_dict[f"Pagina_{num_pag + 1}"] = (
                  estruturar_extrato_banco_brasil(linhas_pag)
              )

        progress_bar.progress(80)
        status_text.text("A gerar a planilha final organizada...")

        df_completo = estruturar_extrato_banco_brasil(todas_linhas_brutas)

        if not df_completo.empty:
          progress_bar.progress(100)
          status_text.empty()
          progress_bar.empty()

          st.success("✨ Processamento e organização concluídos com sucesso!")

          tab1, tab2 = st.tabs(
              ["📋 Pré-visualização", "📥 Central de Download"]
          )

          with tab1:
            st.markdown("### Visualização prévia organizada:")
            st.dataframe(df_completo, use_container_width=True, height=350)

          with tab2:
            st.markdown("### O seu ficheiro está pronto para baixar!")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              if "Tabela única" in modo_export:
                df_completo.to_excel(
                    writer, index=False, sheet_name="Extrato_Limpo"
                )
              else:
                for nome_aba, df_p in paginas_dict.items():
                  df_p.to_excel(writer, index=False, sheet_name=nome_aba)

            excel_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Excel Organizada (.xlsx)",
                data=excel_data,
                file_name="extrato_organizado.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        else:
          progress_bar.empty()
          status_text.empty()
          st.warning("⚠️ Não foi possível estruturar os dados.")

      except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Ocorreu um erro no processamento: {e}")
