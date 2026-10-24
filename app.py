import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor Inteligente de PDF para Excel", page_icon="📊", layout="centered"
)

st.title("📊 Conversor Inteligente de PDF para Excel")
st.write(
    "Envie o seu PDF. O sistema está otimizado para detetar tabelas sem linhas visíveis (como extratos bancários)."
)

uploaded_file = st.file_uploader(
    "Escolha um ficheiro PDF", type=["pdf"], key="pdf"
)

if uploaded_file is not None:
  st.info("A analisar o documento com extração avançada baseada em texto...")

  try:
    with pdfplumber.open(uploaded_file) as pdf:
      all_dfs = []

      # Configuração essencial para documentos sem linhas desenhadas (extratos)
      table_settings = {
          "vertical_strategy": "text",
          "horizontal_strategy": "text",
          "snap_tolerance": 3,
          "join_tolerance": 3,
      }

      for i, page in enumerate(pdf.pages):
        # Extrai tabelas usando o posicionamento do texto
        tables = page.extract_tables(table_settings)

        for table in tables:
          if table and len(table) > 1:
            headers = table[0]
            data = table[1:]

            # Limpeza de cabeçalhos duplicados ou vazios
            clean_headers = []
            seen = {}
            for h in headers:
              h_str = str(h).strip() if h is not None and str(h).strip() != "" else "Coluna"
              h_str = h_str.replace("\n", " ")
              if h_str in seen:
                seen[h_str] += 1
                h_str = f"{h_str}_{seen[h_str]}"
              else:
                seen[h_str] = 0
              clean_headers.append(h_str)

            # Ajusta o tamanho dos dados para coincidir com os cabeçalhos, se necessário
            cleaned_data = []
            for row in data:
              if len(row) == len(clean_headers):
                cleaned_data.append(row)
              elif len(row) < len(clean_headers):
                # Preenche colunas em falta com vazios
                cleaned_data.append(row + [""] * (len(clean_headers) - len(row)))
              else:
                # Corta excessos se houver desfasamento
                cleaned_data.append(row[:len(clean_headers)])

            if cleaned_data:
              df_temp = pd.DataFrame(cleaned_data, columns=clean_headers)
              all_dfs.append(df_temp)

      if all_dfs:
        st.success(f"Foram detetadas e convertidas {len(all_dfs)} tabelas com sucesso!")
        
        # Mostra uma prévia da primeira tabela encontrada
        st.dataframe(all_dfs[0].head(10))

        # Prepara o ficheiro Excel para download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          # Guarda cada tabela numa folha diferente ou junta-as
          for idx, df_t in enumerate(all_dfs):
            sheet_name = f"Tabela_{idx+1}"[:31]  # O Excel limita o nome das folhas a 31 caracteres
            df_t.to_excel(writer, index=False, sheet_name=sheet_name)

        excel_data = output.getvalue()

        st.download_button(
            label="📥 Baixar Planilha Excel (.xlsx)",
            data=excel_data,
            file_name="extrato_bancario_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning(
            "Não foi possível estruturar tabelas automaticamente. Este extrato pode exigir uma leitura linha a linha."
        )

  except Exception as e:
    st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
