import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Conversor Inteligente de PDF para Excel", page_icon="📊", layout="centered"
)

st.title("📊 Conversor Inteligente de PDF para Excel")
st.write(
    "Envie o seu PDF. O sistema vai extrair todas as tabelas estruturadas detetadas no documento."
)

uploaded_file = st.file_uploader(
    "Escolha um ficheiro PDF", type=["pdf"], key="pdf"
)

if uploaded_file is not None:
  st.info("A analisar o documento e a extrair as tabelas...")

  try:
    with pdfplumber.open(uploaded_file) as pdf:
      all_dfs = []

      # Varre todas as páginas do PDF
      for i, page in enumerate(pdf.pages):
        # Extrai TODAS as tabelas da página (no plural)
        tables = page.extract_tables()
        
        for table in tables:
          if table and len(table) > 1:
            headers = table[0]
            data = table[1:]

            # Limpa colunas duplicadas ou vazias para evitar conflitos no Pandas
            clean_headers = []
            seen = {}
            for h in headers:
              h_str = str(h).strip() if h is not None and str(h).strip() != "" else "Coluna"
              # Remove quebras de linha dentro do cabeçalho para o Excel ficar limpo
              h_str = h_str.replace("\n", " ")
              if h_str in seen:
                seen[h_str] += 1
                h_str = f"{h_str}_{seen[h_str]}"
              else:
                seen[h_str] = 0
              clean_headers.append(h_str)

            df_temp = pd.DataFrame(data, columns=clean_headers)
            all_dfs.append(df_temp)

      if all_dfs:
        st.success(f"Foram encontradas {len(all_dfs)} tabelas no documento!")
        
        # Mostra uma prévia da primeira tabela encontrada
        st.dataframe(all_dfs[0].head(10))

        # Prepara o ficheiro Excel (.xlsx) contendo todas as tabelas empilhadas ou numa única estrutura
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          # Se houver várias tabelas, podemos juntá-las ou exportar a principal. 
          # Aqui juntamos as compatíveis ou guardamos a primeira como exemplo principal:
          all_dfs[0].to_excel(writer, index=False, sheet_name="Tabela_Principal")
          
          # Se quiser guardar todas numa folha só ou concatenar:
          combined_df = pd.concat(all_dfs, ignore_index=True, sort=False)
          combined_df.to_excel(writer, index=False, sheet_name="Todas_As_Tabelas")

        excel_data = output.getvalue()

        st.download_button(
            label="📥 Baixar Planilha Excel Completa (.xlsx)",
            data=excel_data,
            file_name="extrato_convertido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
      else:
        st.warning(
            "Não foi encontrada nenhuma tabela estruturada neste PDF."
        )

  except Exception as e:
    st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
