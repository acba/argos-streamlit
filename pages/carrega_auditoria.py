import streamlit as st
import json
from classes import gerar_tabela_encaminhamentos, gerar_tabela_achados, gerar_tabela_situacoes_inconformes, Auditado

st.set_page_config(page_title="Carregar Resultado", layout="wide")

st.title("Carregar Resultado de Auditoria")
st.write("Esta seção permite carregar um resultado de auditoria previamente salvo (arquivo .json) para visualizar e baixar os resultados sem a necessidade de reprocessar os arquivos de entrada.")

arquivo_resultado = st.file_uploader("Carregar arquivo de resultado da auditoria (.json)", type=["json"])

if arquivo_resultado:
    try:
        with st.spinner("Carregando e processando resultado..."):
            # Carrega o objeto 'auditados' do arquivo json
            loaded_data = json.load(arquivo_resultado)
            auditados = {k: Auditado.from_dict(v) for k, v in loaded_data.items()}

            # Gera novamente as tabelas a partir dos dados carregados
            tabela_encaminhamentos = gerar_tabela_encaminhamentos(auditados)
            tabela_achados = gerar_tabela_achados(auditados)
            tabela_situacoes = gerar_tabela_situacoes_inconformes(auditados)

            # Atualiza o estado da sessão para refletir os dados carregados
            st.session_state.audit_results = {
                "auditados": auditados,
                "tabela_encaminhamentos": tabela_encaminhamentos,
                "tabela_achados": tabela_achados,
                "tabela_situacoes": tabela_situacoes,
            }
            st.session_state.audit_completed = True
            st.session_state.files_processed = True # Marca como processado para consistência
            st.session_state.download_files = {} # Limpa arquivos de download antigos
            st.rerun()

    except Exception as e:
        st.error(f"Ocorreu um erro ao carregar o arquivo: {e}")

if st.session_state.audit_completed:
    st.success("Resultado da auditoria carregado com sucesso! Navegue para 'Visualizar Resultado' ou 'Gerar Relatórios'.")