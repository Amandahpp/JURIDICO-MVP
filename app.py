import streamlit as st
import tempfile
import pandas as pd

from extrator import extrair_campos
from gerador import gerar_documento
from buscar_substituir import aplicar_substituicoes, contar_ocorrencias

st.set_page_config(
    page_title="Sistema de Petições Jurídicas",
    page_icon="⚖️",
    layout="wide"
)

st.title("⚖️ Sistema de Geração de Petições")

modo = st.radio(
    "Como o modelo Word está preparado?",
    [
        "Tenho campos marcados com {{ }} no documento",
        "Não quero marcar nada — só trocar um texto que se repete no documento"
    ]
)

st.divider()

arquivo = st.file_uploader(
    "Selecione um modelo Word (.docx)",
    type=["docx"]
)

if arquivo:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(arquivo.read())
        caminho = tmp.name

    # ------------------------------------------------------------------
    # MODO 1: campos marcados com {{ }}
    # ------------------------------------------------------------------
    if modo.startswith("Tenho campos"):

        st.info(
            """
            **Como marcar o modelo no Word:**

            Coloque entre chaves duplas qualquer trecho que precise poder ser
            alterado — pode ser o dado real que já está no documento, por exemplo:

            `{{Amanda}}`, `{{04565447}}`, `{{Rua das Flores, 100}}`

            Ou um nome genérico de campo, se preferir: `{{nome}}`, `{{cpf}}`.

            O sistema mostra o valor já preenchido, prontinho pra você conferir ou
            editar. Se o mesmo texto entre chaves aparecer várias vezes no documento,
            todas as ocorrências mudam juntas quando você editar uma vez.
            """
        )

        campos = extrair_campos(caminho)

        if not campos:
            st.warning(
                "Nenhum campo no formato {{...}} foi encontrado no documento. "
                "Verifique se o modelo usa chaves duplas, ex: {{nome}}."
            )

        else:
            st.success(f"{len(campos)} campo(s) encontrado(s).")
            st.subheader("Confira ou edite os valores")

            dados = {}

            for item in campos:
                dados[item["campo"]] = st.text_input(
                    label=item["campo"],
                    value=item["campo"]
                )

            if st.button("📄 Gerar Documento", key="gerar_modo_campos"):

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

    # ------------------------------------------------------------------
    # MODO 2: buscar e substituir, sem marcadores
    # ------------------------------------------------------------------
    else:

        st.info(
            """
            **Como funciona:**

            Não precisa mexer no Word. Digite abaixo o texto que já está no
            documento (ex: "Amanda Souza") e o texto novo que deve entrar no
            lugar. Todas as vezes que esse texto aparecer no documento serão
            trocadas de uma vez — no início, no meio, na assinatura, onde for.

            Você pode adicionar mais de uma troca por vez, clicando em "+"
            na tabela abaixo.
            """
        )

        tabela_editavel = st.data_editor(
            pd.DataFrame([{"Texto atual no documento": "", "Novo texto": ""}]),
            num_rows="dynamic",
            use_container_width=True,
            key="tabela_substituicoes"
        )

        pares = {}
        for _, linha in tabela_editavel.iterrows():
            original = str(linha.get("Texto atual no documento", "")).strip()
            novo = str(linha.get("Novo texto", "")).strip()
            if original and original != "nan":
                pares[original] = novo

        if pares:
            st.caption("Ocorrências encontradas no documento:")
            for original in pares:
                n = contar_ocorrencias(caminho, original)
                if n == 0:
                    st.warning(f'"{original}" não foi encontrado no documento (verifique acentos/maiúsculas).')
                else:
                    st.write(f'🔎 "{original}" → encontrado {n}x')

        if st.button("📄 Gerar Documento", key="gerar_modo_livre"):

            if not pares:
                st.error("Adicione pelo menos um texto para substituir.")

            elif any(novo.strip() == "" for novo in pares.values()):
                st.error("Preencha o 'Novo texto' para todas as linhas.")

            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as saida_tmp:
                    saida = saida_tmp.name

                aplicar_substituicoes(caminho, pares, saida)

                st.success("Documento gerado com sucesso!")

                with open(saida, "rb") as file:
                    st.download_button(
                        label="⬇️ Baixar Documento",
                        data=file,
                        file_name="peticao.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )