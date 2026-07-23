import streamlit as st
import tempfile

from extrator import extrair_campos
from gerador import gerar_documento

st.set_page_config(
    page_title="Sistema de Petições Jurídicas",
    page_icon="⚖️",
    layout="wide"
)

st.title("⚖️ Sistema de Geração de Petições")
st.write("Envie um modelo Word (.docx) com os campos marcados usando chaves duplas.")

st.info(
    """
    **Como marcar o modelo no Word:**

    Escreva os campos que devem ser preenchidos entre chaves duplas, por exemplo:

    `{{nome}}`, `{{cpf}}`, `{{processo}}`, `{{endereco}}`

    Você pode usar qualquer nome de campo e repetir o mesmo campo várias vezes no
    documento (ex: `{{nome}}` aparecendo em 3 lugares diferentes) — todas as
    ocorrências serão substituídas juntas, pelo mesmo valor.
    """
)

arquivo = st.file_uploader(
    "Selecione um modelo Word (.docx)",
    type=["docx"]
)

if arquivo:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(arquivo.read())
        caminho = tmp.name

    campos = extrair_campos(caminho)

    if not campos:
        st.warning(
            "Nenhum campo no formato {{...}} foi encontrado no documento. "
            "Verifique se o modelo usa chaves duplas, ex: {{nome}}."
        )

    else:
        st.success(f"{len(campos)} campo(s) encontrado(s).")

        st.subheader("Preencha os campos identificados")

        dados = {}

        for item in campos:
            st.caption(f'Marcador no documento: {item["texto_original"]}')

            dados[item["campo"]] = st.text_input(
                label=item["campo"],
                value=""
            )

        if st.button("📄 Gerar Documento"):

            if any(valor.strip() == "" for valor in dados.values()):
                st.error("Preencha todos os campos antes de gerar o documento.")

            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as saida_tmp:
                    saida = saida_tmp.name

                gerar_documento(caminho, dados, saida)

                st.success("Documento gerado com sucesso!")

                with open(saida, "rb") as file:
                    st.download_button(
                        label="⬇️ Baixar Documento",
                        data=file,
                        file_name="peticao.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
