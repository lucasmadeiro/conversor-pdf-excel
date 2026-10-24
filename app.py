import io
import pandas as pd
import pdfplumber
import streamlit as st

# Configuração da página do site
st.set_page_config(
    page_title="Conversor Inteligente de PDF para Excel", page_icon="🤖", layout="centered"
)

st.title("🤖 Conversor Inteligente de PDF para Excel")
st.write(
    "Envie seu PDF. O sistema detectará automaticamente se é uma tabela ou um documento de texto."
)

# Botão para upload do arquivo
uploaded_file = st.file_uploader(
    "Escolha um arquivo PDF", type=["pdf"], key="pdf"
)

if uploaded_file is not None:
  st.info("Analisando o documento e processando os dados...")

  try:
    with pdfplumber.open(uploaded_file) as pdf:
      all_tables = []
      all_texts = []

      # Varre todas as páginas para ver o que o documento contém
      for i, page in enumerate(pdf.pages):
        # Tenta extrair tabela
        table = page.extract_table()
        if table:
          all_tables.append(table)
        
        # Também extrai o texto caso seja um documento corrido
        text = page.extract_text()
        if text:
          lines = text.split("\n")
          for line in lines:
            if line.strip():
              all_texts.append([line.strip()])

      # DECISÃO AUTOMÁTICA: Se encontrou tabelas estruturadas
      if all_tables:
        st.success("📊 Tabela estruturada detectada e extraída com sucesso!")
        
        # Pega a primeira tabela válida como base
        headers = all_tables[0][0]
        data = all_tables[0][1:]
        
        # Limpa colunas duplicadas ou vazias
        clean_headers = []
        seen = {}
        for h in headers:
          h_str = str(h).strip() if h is not None and str(h).strip() != "" else "Coluna"
          if h_str in seen:
            seen[h_str] += 1
            h_str = f"{h_str}_{seen[h_str]}"
          else:
            seen[h_str] = 0
          clean_headers.append(h_str)

        final_df = pd.DataFrame(data, columns=clean_headers)
        
      # CASO CONTRÁRIO: É um documento de texto (como editais)
      elif all_texts:
        st.success("📄 Documento de texto/edital detectado! Convertendo linha por linha...")
        final_df = pd.DataFrame(all_texts, columns=["Conteúdo do Documento"])
        
      else:
        final_df = None

      # Se gerou dados com sucesso, mostra e libera o download
      if final_df is not None and not final_df.empty:
        st.dataframe(final_df.head(10))

        # Prepara o arquivo Excel (.xlsx)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          final_df.to_excel(writer, index=False, sheet_name="Dados_Convertidos")
        excel_data = output.getvalue()

        # Botão de Download
        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=excel_data,
            file_name="arquivo_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning(
            "Não foi possível extrair dados úteis deste PDF. O arquivo pode ser uma imagem escaneada."
        )

  except Exception as e:
    st.error(f"Ocorreu um erro ao processar o arquivo: {e}")
