import io
import re
from pdf2image import convert_from_bytes
import pandas as pd
import pypdf
import pytesseract
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


def extrair_tabela_banco_brasil(file_bytes):
  """Parser dedicado do Banco do Brasil (Mantido 100% intacto)."""
  lancamentos = []
  padrao_data = re.compile(r"^(\d{2}/\d{2}/\d{4})")

  reader = pypdf.PdfReader(io.BytesIO(file_bytes))
  for page in reader.pages:
    texto = page.extract_text()
    if not texto:
      continue
    for linha in texto.split("\n"):
      linha = linha.strip()
      if not linha:
        continue
      if any(
          termo in linha.lower()
          for termo in [
              "extrato conta corrente",
              "cliente conta atual",
              "agência",
              "periodo do extrato",
              "lançamentos",
              "dt. movimento",
              "dt. balancete",
              "transação efetuada",
              "serviço de atendimento",
              "ouvidoria",
              "para deficientes",
              "g3321414432907891",
              "governo",
          ]
      ):
        continue

      match_data = padrao_data.match(linha)
      if match_data:
        data = match_data.group(1)
        resto = linha.replace(data, "", 1).strip()
        valores_c_d = re.findall(r"([\d\.,]+\s*[CD])", resto, re.IGNORECASE)
        valor = ""
        if valores_c_d:
          valor = re.sub(r"[€R\$]", "", valores_c_d[0]).strip()
          for v in valores_c_d:
            resto = resto.replace(v, "").strip()

        partes = resto.split()
        documento = ""
        hist_parts = []
        for parte in partes:
          if (
              re.match(r"^(\d{2}\.\d{3}\.\d{3}|\d{2,15}\.?\d*|\d+)$", parte)
              and not documento
              and len(hist_parts) > 0
          ):
            documento = parte
          else:
            hist_parts.append(parte)

        historico = " ".join(hist_parts).strip()
        if (
            "saldo anterior" in historico.lower()
            or historico.lower() == "saldo"
        ):
          continue

        if historico or valor:
          lancamentos.append({
              "Data": data,
              "Histórico": historico,
              "Documento": documento,
              "Valor (R$)": valor,
          })
      else:
        if lancamentos:
          lancamentos[-1]["Histórico"] += " " + linha

  return (
      pd.DataFrame(lancamentos)
      if lancamentos
      else pd.DataFrame(columns=["Data", "Histórico", "Documento", "Valor (R$)"])
  )


def extrair_itau_consolidado(file_bytes):
  """Parser dedicado e robusto para o Extrato Consolidado (Empresas) do Itaú."""
  lancamentos = []
  padrao_data = re.compile(r"^(\d{2}/\d{2})")

  reader = pypdf.PdfReader(io.BytesIO(file_bytes))
  gravando = False
  ultima_data = ""

  for page in reader.pages:
    texto = page.extract_text()
    if not texto:
      continue
    linhas = texto.split("\n")
    for linha in linhas:
      linha = linha.strip()
      if not linha:
        continue

      # Inicia a gravação a partir do saldo anterior
      if "saldo anterior" in linha.lower():
        gravando = True
        parts = linha.split()
        val = (
            parts[-1]
            if len(parts) > 1 and re.search(r"[\d\.]+,\d{2}", parts[-1])
            else ""
        )
        lancamentos.append({
            "Data": "30/11",
            "Histórico": "Saldo anterior",
            "Valor (R$)": val,
        })
        continue

      # Para a gravação ao atingir os resumos e secções finais
      if any(
          term in linha.lower()
          for term in [
              "conta corrente | aplicações automáticas",
              "resumo - mês",
              "pacote de serviços",
              "notas explicativas",
              "movimentação - aplicações",
          ]
      ):
        gravando = False

      if gravando:
        # Filtra ruídos, cabeçalhos de páginas e notas explicativas de legendas
        if any(
            noise in linha.lower()
            for noise in [
                "extrato mensal",
                "ag 9249",
                "dez 2023",
                "278500",
                "001|",
                "002|",
                "003|",
                "004|",
                "005|",
                "006|",
                "007|",
                "008|",
                "a =",
                "b =",
                "c =",
                "d =",
                "g =",
                "p =",
                "bolsa de valores",
                "crédito a compensar",
                "débito a compensar",
                "este material está disponível",
                "para demais siglas",
                "explicativas",
            ]
        ):
          continue

        match_data = padrao_data.match(linha)
        if match_data:
          ultima_data = match_data.group(1)
          resto = linha.replace(ultima_data, "", 1).strip()
        else:
          resto = linha

        # Se a linha for apenas um valor monetário isolado, atribui ao lançamento anterior
        if re.search(r"^[-]?[\d\.]+,\d{2}-?$", resto):
          if lancamentos:
            lancamentos[-1]["Valor (R$)"] = resto
        else:
          partes = resto.split()
          valor = ""
          hist_partes = []
          for parte in partes:
            if re.search(r"^[-]?[\d\.]+,\d{2}-?$", parte):
              valor = parte
            else:
              hist_partes.append(parte)

          historico = " ".join(hist_partes).strip()
          if historico or valor:
            if any(
                h in historico.lower()
                for h in [
                    "data",
                    "descrição",
                    "entradas",
                    "saídas",
                    "saldo",
                    "total",
                ]
            ):
              continue
            lancamentos.append({
                "Data": ultima_data if ultima_data else "",
                "Histórico": historico,
                "Valor (R$)": valor,
            })

  df = pd.DataFrame(lancamentos)
  # Remove linhas em branco onde o valor ficou vazio
  if not df.empty:
    df = df[df["Valor (R$)"].astype(str).str.strip() != ""]
  return (
      df
      if not df.empty
      else pd.DataFrame(columns=["Data", "Histórico", "Valor (R$)"])
  )


def extrair_itau_movimentacao(file_bytes):
  """Parser dedicado para o Extrato de Movimentação (Conta Corrente) do Itaú."""
  lancamentos = []
  padrao_data = re.compile(r"^(\d{2}/\d{2}/\d{4})")

  reader = pypdf.PdfReader(io.BytesIO(file_bytes))
  for page in reader.pages:
    texto = page.extract_text()
    if not texto:
      continue
    for linha in texto.split("\n"):
      linha = linha.strip()
      if not linha:
        continue

      match_data = padrao_data.match(linha)
      if match_data:
        data = match_data.group(1)
        resto = linha.replace(data, "", 1).strip()

        if "SALDO DO DIA" in resto.upper():
          continue

        partes = resto.split()
        valor = ""
        hist_partes = []

        for parte in partes:
          if re.search(r"^[-]?[\d\.]+,\d{2}$", parte):
            valor = parte
          else:
            hist_partes.append(parte)

        historico = " ".join(hist_partes).strip()

        if historico or valor:
          lancamentos.append({
              "Data": data,
              "Histórico": historico,
              "Valor (R$)": valor,
          })

  return (
      pd.DataFrame(lancamentos)
      if lancamentos
      else pd.DataFrame(columns=["Data", "Histórico", "Valor (R$)"])
  )


def parser_generico(file_bytes):
  """Parser genérico para outros bancos."""
  lancamentos = []
  reader = pypdf.PdfReader(io.BytesIO(file_bytes))
  for page in reader.pages:
    texto = page.extract_text()
    if texto:
      for linha in texto.split("\n"):
        if linha.strip():
          lancamentos.append({
              "Data": "",
              "Histórico": linha.strip(),
              "Valor (R$)": "",
          })
  return pd.DataFrame(lancamentos)


# --- BARRA LATERAL ---
with st.sidebar:
  st.image("logo.png", use_container_width=True)
  st.markdown("---")
  st.title("Painel de Controle")
  st.markdown("---")
  st.markdown("### 💡 Como Funciona")
  st.info(
      "1. Selecione a instituição / modelo.\n2. Envie o extrato em PDF.\n3. Ative"
      " o **OCR** se necessário.\n4. Clique em **Iniciar Processamento**."
  )
  st.markdown("---")
  st.markdown(
      "<p style='text-align: center; font-size: 0.8rem;'>Desenvolvido com"
      " Streamlit & Python</p>",
      unsafe_allow_html=True,
  )

# --- CORPO PRINCIPAL ---
st.title("⚡ Conversor Inteligente de Extratos Bancários")
st.markdown(
    "<p class='upload-text'>Selecione o banco correspondente para aplicar a"
    " regra de leitura ideal.</p>",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="converter-header">
        <div class="badge-file badge-pdf">📄 Extrato PDF</div>
        <div class="arrow-icon">➔ ⚡ ➔</div>
        <div class="badge-file badge-excel">📊 Planilha Excel</div>
    </div>
""",
    unsafe_allow_html=True,
)

col1, col2, col3 = st.columns([1, 6, 1])

with col2:
  banco_selecionado = st.selectbox(
      "🏦 Selecione a Instituição / Modelo:",
      [
          "Banco do Brasil",
          "Itaú - Extrato Consolidado (Empresas)",
          "Itaú - Extrato de Movimentação (Conta Corrente)",
          "Outro / Genérico",
      ],
  )

  uploaded_file = st.file_uploader("", type=["pdf"], key="pdf")

  st.markdown("### Configurações de Extração")

  usar_ocr = st.checkbox(
      "🔍 Ativar OCR (Recomendado para PDFs escaneados ou sem texto"
      " selecionável)",
      value=False,
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

      file_bytes = uploaded_file.getvalue()

      try:
        progress_bar.progress(30)
        status_text.text(f"A processar extrato ({banco_selecionado})...")

        if banco_selecionado == "Banco do Brasil":
          df_completo = extrair_tabela_banco_brasil(file_bytes)
        elif banco_selecionado == "Itaú - Extrato Consolidado (Empresas)":
          df_completo = extrair_itau_consolidado(file_bytes)
        elif banco_selecionado == "Itaú - Extrato de Movimentação (Conta Corrente)":
          df_completo = extrair_itau_movimentacao(file_bytes)
        else:
          df_completo = parser_generico(file_bytes)

        progress_bar.progress(80)
        status_text.text("A estruturar a planilha final...")

        if not df_completo.empty:
          progress_bar.progress(100)
          status_text.empty()
          progress_bar.empty()

          st.success("✨ Processamento concluído com sucesso!")

          tab1, tab2 = st.tabs(
              ["📋 Pré-visualização", "📥 Central de Download"]
          )

          with tab1:
            st.markdown("### Visualização prévia da conversão:")
            st.dataframe(df_completo, use_container_width=True, height=350)

          with tab2:
            st.markdown("### O seu ficheiro está pronto para baixar!")

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
              df_completo.to_excel(
                  writer, index=False, sheet_name="Extrato_Convertido"
              )

            excel_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=excel_data,
                file_name=f"extrato_{banco_selecionado.lower().replace(' ', '_').replace('(', '').replace(')', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
          progress_bar.empty()
          status_text.empty()
          st.warning("⚠️ Nenhum conteúdo extraível encontrado no documento.")

      except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Ocorreu um erro no processamento: {e}")
