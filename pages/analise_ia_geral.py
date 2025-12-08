import os
import tempfile

import streamlit as st
from jinja2 import Environment, BaseLoader, StrictUndefined

from google import genai
from google.genai import types
from utils import avalia_gemini

st.set_page_config(page_title="Análise Geral com IA", layout="wide")

st.title("🤖 Análise Geral com IA (Gemini)")
st.write(
    "Use o poder do Gemini para analisar prompts e arquivos de contexto de forma geral. "
    "Forneça sua chave de API, um prompt detalhado e os arquivos de contexto para obter um resultado estruturado."
)

# --- 1. Configuração da API Key ---
st.subheader("1. Configure sua API Key do Google Gemini")
api_key = st.text_input(
    "GEMINI_API_KEY",
    type="password",
    help="Sua chave não será armazenada. É necessária para cada sessão."
)

if not api_key:
    st.warning("Por favor, insira sua GEMINI_API_KEY para continuar.")
    st.stop()

try:
    # Cria o cliente da API Gemini
    client = genai.Client(api_key=api_key)
except Exception as e:
    st.error(f"Erro ao criar o cliente da API do Gemini. Verifique se a chave é válida. Detalhes: {e}")
    st.stop()

# --- Seleção do Modelo Gemini ---
st.subheader("1.1. Selecione o Modelo Gemini")

model_options = {
    "Gemini 2.5 Flash (gemini-2.5-flash)": "gemini-2.5-flash",
    "Gemini 2.5 Pro (gemini-2.5-pro)": "gemini-2.5-pro",
    "Gemini 2.5 Flash Preview (gemini-2.5-flash-preview-09-2025)": "gemini-2.5-flash-preview-09-2025"
}

selected_model_display = st.selectbox(
    "Escolha o modelo Gemini para a análise:",
    options=list(model_options.keys())
)
selected_model_id = model_options[selected_model_display]

# --- 2. Entrada do Usuário ---
st.subheader("2. Forneça o Prompt e os Arquivos de Contexto")

prompt_template = st.text_area(
    "Escreva seu prompt aqui:",
    height=200,
    placeholder="Exemplo: Resuma o conteúdo dos arquivos fornecidos, destacando os pontos principais. "
                "Retorne o resultado em formato de texto corrido."
)

st.markdown("---")
st.subheader("2.1. Configure a Saída da Análise")

col1, col2 = st.columns(2)
with col1:
    response_format = st.radio(
        "Formato da Resposta:",
        ("Texto", "JSON"),
        horizontal=True,
        help="Selecione JSON para instruir o modelo a retornar uma saída estruturada."
    )
with col2:
    temperature = st.slider(
        "Temperatura:", min_value=0.0, max_value=2.0, value=0.0, step=0.1,
        help="Valores mais baixos geram respostas mais determinísticas. Valores mais altos geram respostas mais criativas."
    )

context_files = st.file_uploader(
    "Carregue seus arquivos de contexto (.txt, .md, .csv, .pdf, etc.)",
    accept_multiple_files=True,
    type=['txt', 'md', 'csv', 'pdf', 'docx'] # Adicionado mais tipos de arquivo para análise geral
)

# --- 3. Geração do Resultado ---
st.subheader("3. Gere a Análise")

# Inicializa o ambiente Jinja2 (mesmo que não use variáveis de auditado, pode ser útil para outras variáveis)
jinja_env = Environment(loader=BaseLoader(), undefined=StrictUndefined)

if st.button("Analisar com Gemini"):
    if not prompt_template:
        st.error("O campo de prompt não pode estar vazio.")
    else:
        with st.spinner("Analisando documentos... Isso pode levar alguns minutos."):
            try:
                # Renderiza o prompt (sem contexto de auditado específico nesta página)
                template = jinja_env.from_string(prompt_template)
                rendered_prompt = template.render()

                st.expander("Prompt Final (clique para expandir)").code(rendered_prompt)

                uploaded_file_objects = []
                if context_files:
                    st.write("Arquivos de contexto carregados para esta análise:")
                    for file in context_files:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=f"_{file.name}") as tmp_file:
                            tmp_file.write(file.getvalue())
                            tmp_file_path = tmp_file.name

                        st.info(f"📄 Fazendo upload de '{file.name}' para a API...")
                        uploaded_file_obj = client.files.create(file=tmp_file_path)
                        uploaded_file_objects.append(uploaded_file_obj)
                        os.remove(tmp_file_path)
                else:
                    st.warning("Nenhum arquivo de contexto carregado para esta análise.")

                # Gera o conteúdo usando o cliente e o modelo selecionado
                response, error_message = avalia_gemini(client, rendered_prompt, selected_model_id, temperature, response_format, uploaded_file_objects)

                if error_message:
                    st.error(error_message)
                elif response:
                    st.session_state.gemini_general_result = response.text
                else:
                    st.error("Nenhuma resposta foi recebida da API do Gemini.")

            except Exception as e:
                st.error(f"Ocorreu um erro durante a chamada para a API do Gemini: {e}")

# --- 4. Exibição do Resultado ---
if 'gemini_general_result' in st.session_state and st.session_state.gemini_general_result:
    st.subheader("Resultado da Análise")
    result_text = st.session_state.gemini_general_result

    # Tenta detectar se o resultado é JSON para uma exibição mais bonita
    if result_text.strip().startswith('{') or result_text.strip().startswith('['):
        st.json(result_text)
    else:
        st.markdown(result_text)

    download_filename = "resultado_gemini_geral.md"
    st.download_button(
        "Baixar Resultado",
        result_text,
        download_filename
    )