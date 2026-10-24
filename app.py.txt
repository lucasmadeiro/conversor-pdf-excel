import io
import pandas as pd
import pdfplumber
import streamlit as st

# Configuração da página do site
st.set_page_config(
    page_title="Conversor de PDF para Excel", page_icon="📊", layout="centered"
)

st.title("📊 Conversor de PDF para Excel")
st.write(
    "Envie seu arquivo PDF abaixo para extrair as tabelas automaticamente para"
    " uma planilha."
)

# Botão para upload do arquivo
uploaded_file = st.file_uploader(
    "Escolha um arquivo PDF", type=["pdf"], key="pdf"
)

if uploaded_file is not None:
  st.info("Lendo o arquivo PDF, por favor aguarde...")

  try:
    # Usando o pdfplumber para abrir e ler o PDF
    with pdfplumber.open(uploaded_file) as pdf:
      all_dfs = []

      # Percorre todas as páginas do PDF procurando tabelas
      for i, page in enumerate(pdf.pages):
        table = page.extract_table()
        if table:
          # Transforma a tabela encontrada em um DataFrame (tabela do Excel)
          df_temp = pd.DataFrame(table[1:], columns=table[0])
          all_dfs.append(df_temp)

      if all_dfs:
        # Junta todas as tabelas encontradas em uma só
        final_df = pd.concat(all_dfs, ignore_index=True)

        st.success("Tabela extraída com sucesso!")
        st.dataframe(
            final_df.head(10)
        )  # Mostra uma prévia das primeiras linhas

        # Prepara o arquivo Excel (.xlsx) para download usando a biblioteca io BytesIO
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          final_df.to_excel(writer, index=False, sheet_name="Dados_PDF")
        excel_data = output.getvalue()

        # Botão de Download do Excel
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=excel_data,
            file_name="tabela_convertida.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      else:
        st.warning(
            "Não foi encontrada nenhuma tabela estruturada neste PDF. Lembre-se"
            " que PDFs gerados de imagens (scans) exigem OCR avançado."
        )

  except Exception as e:
    st.error(f"Ocorreu um erro ao processar o arquivo: {e}")