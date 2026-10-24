import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor de Extratos PDF para Excel", page_icon="📄", layout="centered"
)

st.title("📄 Conversor de Extratos e Documentos PDF")
st.write(
    "Envie o seu PDF. O sistema extrairá o conteúdo de forma limpa e organizada por linhas, ideal para extratos bancários e editais."
)

uploaded_file = st.file_uploader(
    "Escolha um ficheiro PDF", type=["pdf"], key="pdf"
)

if uploaded_file is not None:
  st.info("A processar o documento linha a linha...")

  try:
    with pdfplumber.open(uploaded_file) as pdf:
      linhas_extraidas = []

      # Varre todas as páginas do PDF recolhendo o texto organizado
      for num_pag, page in enumerate(pdf.pages):
        linhas_extraidas.append([f"--- PÁGINA {num_pag + 1} ---"])
        
        texto_pagina = page.extract_text(layout=True)
        if texto_pagina:
          for linha in texto_pagina.split("\n"):
            if linha.strip():
              linhas_extraidas.append([linha.strip()])

      if linhas_extraidas:
        # Cria um DataFrame onde cada linha do PDF vira uma linha na Coluna A do Excel
        df = pd.DataFrame(linhas_extraidas, columns=["Conteúdo do Extrato"])

        st.success("Documento convertido com sucesso!")
        st.dataframe(df.head(15))

        # Prepara o ficheiro Excel para download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df.to_excel(writer, index=False, sheet_name="Extrato_Limpo")

        excel_data = output.getvalue()

        st.download_button(
            label="📥 Baixar Planilha Excel Limpa (.xlsx)",
            data=excel_data,
            file_name="extrato_organizado.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning("Não foi possível extrair texto legível deste PDF.")

  except Exception as e:
    st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
