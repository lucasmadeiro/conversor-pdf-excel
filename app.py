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


def parser_banco_brasil(linhas_texto):
  """Parser cirúrgico para extratos do Banco do Brasil com alinhamento correto de Valor e Saldo."""
  lancamentos = []
  padrao_data = re.compile(r"^(\d{2}/\d{2}/\d{4})")
  
  # Regex para capturar valores monetários com C ou D no final (ex: 6.042,58 C, 0,00 C)
  # Aceita opcionalmente o símbolo de moeda (€ ou R$) se houver
  padrao_valor = re.compile(r"([€R\$\s]*[\d\.,]+\s*[CD])", re.IGNORECASE)

  ultima_data = ""

  for linha in linhas_texto:
    linha = linha.strip()
    if not linha:
      continue

    # Ignora cabeçalhos e rodapés institucionais do BB
    if any(
        termo in linha
        for termo in [
            "GOVERNO",
            "Extrato conta corrente",
            "Cliente Conta atual",
            "Agência",
            "Periodo do extrato",
            "Lançamentos",
            "Dt. movimento",
            "Dt. balancete",
            "Transação efetuada",
            "Serviço de Atendimento",
            "Ouvidoria",
            "Para deficientes",
            "G3321414432907891",
        ]
    ):
      continue

    match_data = padrao_data.match(linha)
    if match_data:
      ultima_data = match_data.group(1)
      resto = linha.replace(ultima_data, "", 1).strip()

      # Extrai todos os valores monetários encontrados na linha
      valores = padrao_valor.findall(resto)
      
      valor = ""
      saldo = ""

      # Atribuição precisa da direita para a esquerda
      if len(valores) >= 2:
        valor = valores[-2].strip()
        saldo = valores[-1].strip()
        # Remove os valores extraídos do resto para limpar o histórico
        for v in valores:
          resto = resto.replace(v, "").strip()
      elif len(valores) == 1:
        val_encontrado = valores[0].strip()
        resto = resto.replace(val_encontrado, "").strip()
        if "saldo" in resto.lower() or "anterior" in resto.lower() or "salto" in resto.lower():
          saldo = val_encontrado
        else:
          valor = val_encontrado

      # Limpa símbolos de moeda extras dos valores
      valor = re.sub(r"[€R\$]", "", valor).strip()
      saldo = re.sub(r"[€R\$]", "", saldo).strip()

      # Identifica o número do documento/lote (números isolados típicos do BB)
      partes_resto = resto.split()
      documento = ""
      historico_partes = []

      for parte in partes_resto:
        if re.match(r"^(\d{2}\.\d{3}\.\d{3}|\d{2,15}\.?\d*|\d+)$", parte) and not documento and len(historico_partes) > 0:
          documento = parte
        else:
          historico_partes.append(parte)

      historico = " ".join(historico_partes).strip()

      lancamentos.append({
          "Data": ultima_data,
          "Histórico": historico if historico else "",
          "Documento": documento if documento else "",
          "Valor (R$)": valor if valor else "",
          "Saldo (R$)": saldo if saldo else "",
      })
    else:
      # Linha de continuação de histórico (ex: nome do favorecido)
      if lancamentos and ultima_data:
        lancamentos[-1]["Histórico"] += " " + linha
      else:
        lancamentos.append({
            "Data": "",
            "Histórico": linha,
            "Documento": "",
            "Valor (R$)": "",
            "Saldo (R$)": "",
        })

  return (
      pd.DataFrame(lancamentos)
      if lancamentos
      else pd.DataFrame({"Conteúdo Bruto": linhas_texto})
  )


def parser_itau(linhas_texto):
  """Parser especializado para extratos do Itaú."""
  lancamentos = []
  padrao_data = re.compile(r"^(\d{2}/\d{2})\b")
  gravando = False
  ultima_data = ""

  for linha in linhas_texto:
    linha = linha.strip()
    if not linha:
      continue

    linha_lower = linha.lower()

    if any(
        parada in linha_lower
        for parada in [
            "principais siglas",
            "notas explicativas",
            "sac 0800",
            "ouvidoria",
            "central de atendimento",
            "este material está disponível",
        ]
    ):
      gravando = False
      continue

    if "saldo anterior" in linha_lower or "lançamentos" in linha_lower:
      gravando = True
      continue

    if not gravando and padrao_data.match(linha):
      gravando = True

    if gravando:
      match_data = padrao_data.match(linha)
      if match_data:
        ultima_data = match_data.group(1)
        conteudo = linha[match_data.end() :].strip()
        conteudo = re.sub(r"-\d+$", "", conteudo).strip()
        lancamentos.append({
            "Data": ultima_data,
            "Histórico": conteudo,
            "Documento": "",
            "Valor (R$)": "",
            "Saldo (R$)": "",
        })
      else:
        conteudo_limpo = re.sub(r"-\d+$", "", linha).strip()
        if lancamentos:
          lancamentos[-1]["Histórico"] += " " + conteudo_limpo
        else:
          lancamentos.append({
              "Data": ultima_data,
              "Histórico": conteudo_limpo,
              "Documento": "",
              "Valor (R$)": "",
              "Saldo (R$)": "",
          })

  return (
      pd.DataFrame(lancamentos)
      if lancamentos
      else pd.DataFrame({"Conteúdo Bruto": linhas_texto})
  )


def parser_generico(linhas_texto):
  """Parser genérico para outros bancos."""
  lancamentos = [
      {"Data": "", "Histórico": l, "Documento": "", "Valor (R$)": "", "Saldo (R$)": ""} 
      for l in linhas_texto if l.strip()
  ]
  return pd.DataFrame(lancamentos)


# --- BARRA LATERAL ---
with st.sidebar:
  st.image("logo.png", use_container_width=True)
  st.markdown("---")
  st.title("Painel de Controle")
  st.markdown("---")
  st.markdown("### 💡 Como Funciona")
  st.info(
      "1. Selecione a instituição financeira.\n2. Envie o extrato em PDF.\n3."
      " Ative o **OCR** se necessário.\n4. Clique em **Iniciar Processamento**."
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
      "🏦 Selecione a Instituição Financeira:",
      ["Banco do Brasil", "Itaú", "Outro / Genérico"],
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

      paginas_dict = {}
      todas_linhas_brutas = []
      file_bytes = uploaded_file.getvalue()

      try:
        if usar_ocr:
          status_text.text("A realizar leitura visual OCR...")
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

            if banco_selecionado == "Banco do Brasil":
              paginas_dict[f"Pagina_{num_pag + 1}"] = parser_banco_brasil(
                  linhas_pag
              )
            elif banco_selecionado == "Itaú":
              paginas_dict[f"Pagina_{num_pag + 1}"] = parser_itau(linhas_pag)
            else:
              paginas_dict[f"Pagina_{num_pag + 1}"] = parser_generico(
                  linhas_pag
              )
        else:
          status_text.text("A ler a camada de texto do PDF...")
          progress_bar.progress(30)

          with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for num_pag, page in enumerate(pdf.pages):
              texto_pagina = page.extract_text(layout=False)
              linhas_pag = []
              if texto_pagina:
                for linha in texto_pagina.split("\n"):
                  if linha.strip():
                    todas_linhas_brutas.append(linha.strip())
                    linhas_pag.append(linha.strip())

              if banco_selecionado == "Banco do Brasil":
                paginas_dict[f"Pagina_{num_pag + 1}"] = parser_banco_brasil(
                    linhas_pag
                )
              elif banco_selecionado == "Itaú":
                paginas_dict[f"Pagina_{num_pag + 1}"] = parser_itau(linhas_pag)
              else:
                paginas_dict[f"Pagina_{num_pag + 1}"] = parser_generico(
                    linhas_pag
                )

        progress_bar.progress(80)
        status_text.text("A estruturar a planilha final...")

        if banco_selecionado == "Banco do Brasil":
          df_completo = parser_banco_brasil(todas_linhas_brutas)
        elif banco_selecionado == "Itaú":
          df_completo = parser_itau(linhas_pag)
        else:
          df_completo = parser_generico(todas_linhas_brutas)

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
              if "Tabela única" in modo_export:
                df_completo.to_excel(
                    writer, index=False, sheet_name="Extrato_Convertido"
                )
              else:
                for nome_aba, df_p in paginas_dict.items():
                  df_p.to_excel(writer, index=False, sheet_name=nome_aba)

            excel_data = output.getvalue()

            st.download_button(
                label="📥 Baixar Planilha Excel (.xlsx)",
                data=excel_data,
                file_name=f"extrato_{banco_selecionado.lower().replace(' ', '_')}.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                ),
            )
        else:
          progress_bar.empty()
          status_text.empty()
          st.warning("⚠️ Nenhum conteúdo extraível encontrado no documento.")

      except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ Ocorreu um erro no processamento: {e}")
