import streamlit as st
import os
import zipfile
import io
import tempfile
from datetime import date
from docxtpl import DocxTemplate
from classes import Auditado

# Configuração da página
st.set_page_config(page_title="Gerar Anexos de Comentários do Gestor", layout="wide")

st.title("Gerar Anexos de Comentários do Gestor")
st.write("Esta ferramenta gera arquivos .docx individuais para cada auditado, contendo os achados e situações encontradas, para que o gestor possa apresentar suas considerações.")

if "audit_completed" in st.session_state and st.session_state.audit_completed:
    results = st.session_state.audit_results
    auditados = results["auditados"]

    # Filtra apenas auditados com achados
    auditados_com_achados = {k: v for k, v in auditados.items() if v.tem_achados}

    st.write(f"Total de auditados processados: **{len(auditados)}**")
    st.write(f"Auditados com achados (elegíveis para documento): **{len(auditados_com_achados)}**")

    if not auditados_com_achados:
        st.warning("Nenhum auditado possui achados para gerar o documento.")
    else:
        # Inputs do usuário
        col1, col2 = st.columns(2)
        with col1:
            data_entrega = st.date_input("Data Final de Entrega:", value=date.today())
            # Formata a data para string (ex: 27/11/2025)
            data_final_entrega = data_entrega.strftime("%d/%m/%Y")

        with col2:
            email_contato = st.text_input("Email de Contato:", value="auditoria_seginfo@tcerj.tc.br")

        # Seleção de auditados
        opcoes_auditados = list(auditados_com_achados.keys())
        selecionados = st.multiselect(
            "Selecione os auditados para gerar o documento (deixe vazio para selecionar todos):",
            options=opcoes_auditados,
            default=opcoes_auditados
        )

        if st.button("Gerar Documentos .docx"):
            # Se nenhum selecionado, usa todos
            siglas_selecionadas = selecionados if selecionados else opcoes_auditados

            template_path = "docs/template-questionario-comentarios-gestor.docx"

            if not os.path.exists(template_path):
                st.error(f"Template não encontrado em: {template_path}")
            else:
                zip_buffer = io.BytesIO()
                files_generated = 0

                with tempfile.TemporaryDirectory() as tmp_dir:
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:

                        progress_bar = st.progress(0)
                        for i, sigla in enumerate(siglas_selecionadas):
                            auditado = auditados_com_achados[sigla]

                            try:
                                doc = DocxTemplate(template_path)

                                # Prepara o contexto
                                achados = [p.achado for p in auditado.procedimentos_executados if p.achado is not None]

                                contexto = {
                                    'auditado': auditado,
                                    'achados': achados,
                                    'data_final_entrega': data_final_entrega,
                                    'email_contato': email_contato
                                }

                                doc.render(contexto)

                                # Salva arquivo temporário
                                filename = f"Anexo - Questionário Comentarios ({sigla}).docx"
                                file_path = os.path.join(tmp_dir, filename)
                                doc.save(file_path)

                                # Adiciona ao ZIP
                                zip_file.write(file_path, arcname=filename)
                                files_generated += 1

                            except Exception as e:
                                st.error(f"Erro ao gerar documento para {sigla}: {e}")

                            progress_bar.progress((i + 1) / len(siglas_selecionadas))

                if files_generated > 0:
                    st.success(f"{files_generated} documentos gerados com sucesso!")
                    st.download_button(
                        label="Baixar Documentos (.zip)",
                        data=zip_buffer.getvalue(),
                        file_name="documentos_comentarios_gestor.zip",
                        mime="application/zip",
                        type="primary"
                    )
                else:
                    st.warning("Nenhum documento foi gerado.")

else:
    st.info("Por favor, carregue ou processe a auditoria primeiro na página 'Carregar Auditoria' ou 'Aplica Procedimentos'.")
