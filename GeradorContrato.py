import streamlit as st
import pandas as pd
from docx import Document
import io

st.set_page_config(page_title="Gerador de Contratos", layout="centered")

st.title("📝 Gerador de Contratos Automático")

st.write("Faça upload de um modelo Word com variáveis no formato `{variavel}` e de uma planilha Excel com duas colunas: `Variável` e `Valor`.")

# Upload dos arquivos
modelo_docx = st.file_uploader("📄 Upload do modelo .docx", type=["docx"])
planilha_excel = st.file_uploader("📊 Upload da planilha .xlsx", type=["xlsx"])

if modelo_docx and planilha_excel:
    try:
        # Carrega a planilha
        df = pd.read_excel(planilha_excel, sheet_name=0)
        variaveis = pd.Series(df["Valor"].values, index=df["Variável"]).to_dict()

        # Carrega o modelo
        doc = Document(modelo_docx)

        # Substitui no texto
        for paragrafo in doc.paragraphs:
            for chave, valor in variaveis.items():
                paragrafo.text = paragrafo.text.replace(f"{{{chave}}}", str(valor))

        # Substitui nas tabelas
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for chave, valor in variaveis.items():
                        celula.text = celula.text.replace(f"{{{chave}}}", str(valor))

        # Prepara para download
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        st.success("✅ Contrato gerado com sucesso!")
        st.download_button(
            label="📥 Baixar Contrato Preenchido",
            data=output,
            file_name="ContratoPreenchido.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
