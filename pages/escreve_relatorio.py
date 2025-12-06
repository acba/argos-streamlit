import streamlit as st
import pandas as pd
import ast
import re
import base64
import textwrap
import tempfile
import os
import logging
import pypandoc
from jinja2 import Environment, BaseLoader, StrictUndefined, exceptions
from streamlit_ace import st_ace

# Import from local modules
from utils import (
    data_hoje_abnt, data_hoje, processa_imagens_contexto, cross_ref_figuras,
    cross_ref_tabelas, substituir_underline_preview, substituir_underline_pandoc,
    StreamlitLogHandler, aplicar_estilo_tabelas, processar_quebras_pagina
)

# --- Configuração da Página ---
st.set_page_config(page_title="Escreve Relatório", layout="wide")

EDITOR_HEIGHT = None # pixels
PREVIEW_HEIGHT = None


# --- Gerenciamento de Estado (Persistência) ---
if 'escreve_relatorio_contexto' not in st.session_state:
    st.session_state.escreve_relatorio_contexto = {}
if 'escreve_relatorio_content' not in st.session_state:
    st.session_state.escreve_relatorio_content = "## Relatório de Auditoria\n\nEscreva seu relatório aqui..."
if 'escreve_relatorio_uploaded_files_keys' not in st.session_state:
    st.session_state.escreve_relatorio_uploaded_files_keys = [] # Rastreia arquivos para limpar se necessário
if 'escreve_relatorio_layout_split' not in st.session_state:
    st.session_state.escreve_relatorio_layout_split = 50 # Default 50%
if 'escreve_relatorio_docx_download' not in st.session_state:
    st.session_state.escreve_relatorio_docx_download = None

# --- Funções Auxiliares ---
def reset_page():
    """Reseta o estado da página."""
    st.session_state.escreve_relatorio_contexto = {}
    st.session_state.escreve_relatorio_content = "## Relatório de Auditoria\n\nEscreva seu relatório aqui..."
    st.session_state.escreve_relatorio_layout_split = 50
    st.session_state.escreve_relatorio_docx_download = None
    st.rerun()

def load_template(file):
    if file is not None:
        stringio = file.getvalue().decode("utf-8")
        st.session_state.escreve_relatorio_content = stringio
        st.rerun() # Rerun to update the editor

def process_images_for_preview(markdown_text, mapa_imagens):
    """
    Substitui as tags de imagem do markdown por tags HTML <img> com base64
    para exibição no preview, considerando a legenda extraída por cross_ref_figuras.
    """

    def img_to_base64(uploaded_file):
        try:
            # UploadedFile se comporta como arquivo aberto, resetar ponteiro se necessário
            uploaded_file.seek(0)
            return base64.b64encode(uploaded_file.read()).decode()
        except Exception:
            return ""

    def replacement(match):
        alt_text = match.group(1) # Pega o alt text, que pode conter "Figura X - Título"
        image_name = match.group(2)

        if image_name in mapa_imagens:
            b64_string = img_to_base64(mapa_imagens[image_name])
            file_ext = image_name.split('.')[-1].lower()
            mime_type = f"image/{file_ext}"
            if file_ext == 'jpg': mime_type = "image/jpeg"
            elif file_ext == 'svg': mime_type = "image/svg+xml"

            # Monta o HTML conforme solicitado:
            # Figura XX - Título
            # [FIGURA]

            html = textwrap.dedent(f"""
            <div style="text-align: center;">
                <p><b>{alt_text}</b></p>
                <img src="data:{mime_type};base64,{b64_string}" alt="{alt_text}" style="max-width: 100%;">
            </div>
            """)
            return html
        else:
            # Se a imagem não está no contexto, mantém ou avisa
            return f"![{alt_text}]({image_name}) <!-- Imagem não encontrada no contexto -->"

    # Regex para capturar a imagem processada pelo cross_ref_figuras
    # Ex: ![Figura 1 - Título](nome_arquivo.png){width=...}
    # ou apenas ![Figura 1 - Título](nome_arquivo.png)
    regex = r"!\[([^\]]*)\]\(([^)]*)\)(?:\{([^}]*)\})?(?:\{#?([^}]*)\})?"

    return re.sub(regex, replacement, markdown_text)

# --- Interface ---

# Botão de Reset no topo
col_header, col_reset = st.columns([5, 1])
with col_header:
    st.title("Escreve Relatório")
with col_reset:
    if st.button("Zerar Tudo", type="primary"):
        reset_page()

st.markdown("Desenvolva seu relatório e verifique o preview em tempo real.")

# --- 1. Seção de Contexto ---
with st.expander("1. Contexto Adicional"):

    with st.expander("Como funciona o Contexto Adicional?"):
        st.markdown("""
        Esta seção permite carregar dados que ficarão disponíveis para uso no seu relatório (template).

        **Tipos de arquivos suportados:**
        - **Planilhas Excel (.xlsx):**
            - Devem ter uma coluna chamada **`sigla`** para vincular os dados.
            - As colunas se tornam variáveis no template (ex: coluna `valor_total` -> `{{ valor_total }}`).
            - Colunas terminadas em `*` (ex: `lista_itens*`) são interpretadas como listas/dicionários Python.
        - **Imagens:**
            - Podem ser referenciadas no template pelo nome do arquivo.

        **Dados de Auditoria do Sistema:**
        - Se você já executou procedimentos de auditoria no sistema, pode selecionar um auditado para trazer seus dados (achados, procedimentos, etc.) para o contexto automaticamente.
        """)

    # 1.1 Seleção de Auditado (do Sistema)
    auditado_context = {}
    if 'audit_completed' in st.session_state and st.session_state.audit_completed:
        results = st.session_state.audit_results
        auditados = results["auditados"]

        # Criar lista de opções
        opcoes_auditados = list(auditados.keys())

        selected_auditado_key = st.selectbox(
            "Adicionar dados de auditoria processada (Opcional):",
            options=["Nenhum"] + opcoes_auditados,
            index=0
        )

        if selected_auditado_key != "Nenhum":
            auditado_obj = auditados[selected_auditado_key]
            # Adiciona ao contexto como 'auditado' e também a sigla para merge com planilhas
            auditado_context['auditado'] = auditado_obj
            auditado_context['sigla'] = auditado_obj.sigla
    else:
        st.info("Nenhum resultado de auditoria carregado no sistema. Você pode carregar manualmente na página 'Carrega Auditoria' ou 'Aplica Procedimentos'.")

    # 1.2 Upload de Arquivos de Contexto
    arquivos_contexto = st.file_uploader("Carregar Planilhas de Contexto (.xlsx)", type=["xlsx"], accept_multiple_files=True, key="upl_contexto")
    df_contexto_extra = None

    if arquivos_contexto:
        dfs_contexto = []
        for arq in arquivos_contexto:
            try:
                df_temp = pd.read_excel(arq)
                if 'sigla' in df_temp.columns:
                    df_temp = df_temp.set_index('sigla')
                    df_temp.columns = [col.strip() for col in df_temp.columns]

                    # Processa colunas especiais terminadas em '*'
                    for col in df_temp.columns:
                        if col.endswith('*'):
                            try:
                                df_temp[col] = df_temp[col].apply(lambda x: ast.literal_eval(x) if isinstance(x, str) else x)
                                df_temp.rename(columns={col: col.rstrip('*')}, inplace=True)
                            except Exception as e:
                                st.warning(f"Falha ao converter coluna especial '{col}' no arquivo '{arq.name}': {e}")
                    dfs_contexto.append(df_temp)
                else:
                    st.warning(f"Arquivo '{arq.name}' ignorado: Coluna 'sigla' não encontrada.")
            except Exception as e:
                st.error(f"Erro ao ler '{arq.name}': {e}")

        if dfs_contexto:
            df_contexto_extra = pd.concat(dfs_contexto).groupby(level=0).first()
            st.success(f"{len(dfs_contexto)} planilhas processadas.")

    # 1.3 Upload de Imagens
    arquivos_imagens = st.file_uploader("Carregar Imagens para o Contexto", accept_multiple_files=True, type=['png', 'jpg', 'jpeg'], key="upl_imgs")
    mapa_imagens = {}
    if arquivos_imagens:
        for img in arquivos_imagens:
            mapa_imagens[img.name] = img
        st.success(f"{len(arquivos_imagens)} imagens carregadas.")


    # --- Montagem do Contexto Final ---
    # O contexto é recriado a cada rerun para garantir consistência com os inputs atuais
    final_context = {}

    # Adiciona helpers globais
    final_context['data_hoje_abnt'] = data_hoje_abnt()
    final_context['data_hoje'] = data_hoje()

    # 1. Dados do Auditado Selecionado (Base)
    if auditado_context:
        final_context.update(auditado_context)

    # 2. Dados das Planilhas (Merge pela Sigla)
    if df_contexto_extra is not None:
        sigla_atual = final_context.get('sigla')

        # Se temos um auditado selecionado, tentamos pegar a linha dele nas planilhas
        if sigla_atual and sigla_atual in df_contexto_extra.index:
            dados_extra = df_contexto_extra.loc[sigla_atual].to_dict()
            final_context.update(dados_extra)
        elif not sigla_atual and len(df_contexto_extra) > 0:
             # Se não tem auditado selecionado, mas tem planilha, avisar ou pegar o primeiro?
             # O requisito diz: "selecionar qual auditado ele deseja adicionar a variavel 'auditado'"
             # Como o df tem index 'sigla', precisamos saber qual linha usar.
             # Vamos permitir selecionar uma linha da planilha se nenhum auditado do sistema estiver selecionado.
             st.info("Nenhum auditado do sistema selecionado. Usando dados da planilha.")
             sigla_planilha = st.selectbox("Selecione a Sigla da planilha para usar no contexto:", df_contexto_extra.index)
             if sigla_planilha:
                 final_context.update(df_contexto_extra.loc[sigla_planilha].to_dict())
                 final_context['sigla'] = sigla_planilha

    # 3. Imagens - Apenas o nome para a renderização do Jinja ficar correta (path relativo)
    for img_name in mapa_imagens:
        final_context[img_name] = img_name

    # Exibir contexto atual (Debug opcional)
    with st.expander("Ver Variáveis de Contexto Disponíveis"):
        st.json({k: str(v)[:200] for k, v in final_context.items()})


# --- 2. Seção de Template ---
with st.expander("2. Template"):
    col_temp1, col_temp2 = st.columns(2)
    with col_temp1:
        # Upload de Template Opcional
        uploaded_template = st.file_uploader("Carregar Conteúdo Markdown (.md)", type=["md"], key="upl_template_content")
        if uploaded_template:
            if st.button("Substituir Conteúdo do Editor pelo Arquivo"):
                load_template(uploaded_template)
    with col_temp2:
        # Upload de Template DOCX (Estilos)
        arquivo_template_docx = st.file_uploader("Template de Estilos (.docx) para geração", type=["docx"], key="upl_template_docx", help="Se não fornecido, usará o padrão do sistema.")


# --- 3. Editor e Preview (Split View) ---

# Área de Ações e Controle
col_btn_gerar, col_ctrl = st.columns([2, 8])

with col_btn_gerar:
    if st.button("Gerar Relatório (.docx)", type="primary"):
        if not st.session_state.escreve_relatorio_content.strip():
            st.error("O editor está vazio.")
        else:
            with st.spinner("Gerando relatório DOCX..."):
                try:
                    # Configuração do Ambiente Jinja2
                    env = Environment(loader=BaseLoader(), undefined=StrictUndefined)

                    # Preparação do Template de Referência
                    template_ref_path = 'docs/template-relatorio-individual.docx'

                    with tempfile.TemporaryDirectory() as tmp_dir:
                        # 1. Salvar Template de Estilos (se houver upload ou usar padrão)
                        if arquivo_template_docx:
                            template_ref_path = os.path.join(tmp_dir, "template_ref.docx")
                            with open(template_ref_path, "wb") as f:
                                f.write(arquivo_template_docx.getvalue())

                        # 2. Salvar Imagens do Contexto no Temp
                        for img_name, img_obj in mapa_imagens.items():
                            img_path = os.path.join(tmp_dir, img_name)
                            img_obj.seek(0)
                            with open(img_path, "wb") as f:
                                f.write(img_obj.read())

                        # 3. Processamento do Markdown (Jinja + Pandoc steps)
                        # Renderiza Jinja
                        template = env.from_string(st.session_state.escreve_relatorio_content)

                        # Para o Pandoc, precisamos que o contexto tenha os CAMINHOS das imagens no temp, não os objetos ou apenas nomes
                        # Criamos um contexto local atualizado com os caminhos absolutos do temp
                        contexto_pandoc = final_context.copy()
                        for img_name in mapa_imagens:
                            contexto_pandoc[img_name] = os.path.join(tmp_dir, img_name)

                        # Processa imagens do contexto (substitui objetos por caminhos se necessário - aqui já fizemos acima)
                        # Mas precisamos garantir que o template use os caminhos corretos.
                        # A função processa_imagens_contexto faria isso se tivéssemos mapeado paths.
                        # Como sobrescrevemos manualmente acima, o render já deve pegar os paths.

                        rendered_markdown = template.render(contexto_pandoc)

                        # Pós-processamento Markdown para Pandoc
                        rendered_markdown = cross_ref_figuras(rendered_markdown)
                        rendered_markdown = cross_ref_tabelas(rendered_markdown)
                        rendered_markdown = processar_quebras_pagina(rendered_markdown)
                        rendered_markdown = substituir_underline_pandoc(rendered_markdown)

                        # Salvar MD temporário
                        md_filename = os.path.join(tmp_dir, 'relatorio_temp.md')
                        with open(md_filename, 'w', encoding='utf-8') as f:
                            f.write(rendered_markdown)

                        # 4. Conversão Pandoc
                        docx_filename = os.path.join(tmp_dir, 'relatorio_final.docx')

                        # Argumentos Pandoc
                        resource_path_arg = '--resource-path=' + tmp_dir
                        args_docx = ['--figure-caption-position=above', '--reference-doc=' + template_ref_path, resource_path_arg]

                        # Logging
                        pypandoc_logger = logging.getLogger('pypandoc')
                        pypandoc_logger.setLevel(logging.WARNING)
                        log_container = st.empty()
                        handler = StreamlitLogHandler(log_container)
                        pypandoc_logger.addHandler(handler)

                        pypandoc.convert_file(md_filename, to='docx', outputfile=docx_filename, extra_args=args_docx)
                        pypandoc_logger.removeHandler(handler)

                        # 5. Estilos de Tabela
                        aplicar_estilo_tabelas(docx_filename)

                        # 6. Ler arquivo final para download
                        with open(docx_filename, "rb") as f:
                            st.session_state.escreve_relatorio_docx_download = f.read()

                    st.success("Relatório gerado com sucesso!")

                except exceptions.UndefinedError as e:
                    st.error(f"Erro de Variável Indefinida: {e}")
                except Exception as e:
                    st.error(f"Erro ao gerar relatório: {e}")

    # Botão de Download (aparece se gerado)
    if st.session_state.escreve_relatorio_docx_download:
        st.download_button(
            label="Baixar Relatório (.docx)",
            data=st.session_state.escreve_relatorio_docx_download,
            file_name=f"relatorio_audit_{data_hoje().replace('/','-')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

with col_ctrl:
    split_value = st.slider("Ajustar largura do Editor (%)", min_value=10, max_value=90, value=st.session_state.escreve_relatorio_layout_split, step=5)
    st.session_state.escreve_relatorio_layout_split = split_value

col_editor, col_preview = st.columns([split_value, 100 - split_value], border=True)

with col_editor:
    st.write("### Editor (Markdown/Jinja)")
    # Editor Ace
    content = st_ace(
        value=st.session_state.escreve_relatorio_content,
        language="markdown",
        theme="monokai",
        key="ace_editor",
        height=EDITOR_HEIGHT,
        # height=600,
        auto_update=True # Atualiza conforme digita (pode ser lento se o template for gigante)
    )
    # Atualiza o estado para persistência
    if content != st.session_state.escreve_relatorio_content:
        st.session_state.escreve_relatorio_content = content
        st.rerun() # Força atualização do preview

with col_preview:
    st.write("### Preview")

    # Custom CSS
    st.markdown(
        textwrap.dedent(f'''
        <style>
            .scrollable-preview{{
                max-height: {PREVIEW_HEIGHT}px;
                overflow-y:auto;
                border: 1px solid #ccc;
                padding: 10px;
                border-radius: 5px;
            }}
        </style>
        '''),
        unsafe_allow_html=True
    )

    if content:
        try:
            # Renderização Jinja2
            env = Environment(loader=BaseLoader(), undefined=StrictUndefined)

            # Pré-processamento simples (similar ao gera_relatorios)
            # Nota: Não estamos processando includes de arquivos externos aqui no preview tempo real para simplicidade

            try:
                template_md = cross_ref_figuras(content)
                template_md = cross_ref_tabelas(template_md)
                template_md = substituir_underline_preview(template_md)

            except Exception as e:
                st.warning(f"Aviso no pre-processamento: {e}")

            template_md = env.from_string(template_md)
            template_md_processado = template_md.render(final_context)

            # Pós-processamento de visualização
            # Cross-ref apenas se tiver func
            try:
                # Processa Imagens para o Preview (Base64)
                template_md_processado = process_images_for_preview(template_md_processado, mapa_imagens)

            except Exception as e:
                # pass # Ignora erros de pós-processamento no preview
                st.warning(f"Aviso no pós-processamento: {e}")

            # st.markdown(template_md_processado, unsafe_allow_html=True)
            st.markdown(f'<div class="scrollable-preview">{template_md_processado}</div>', unsafe_allow_html=True)

        except exceptions.UndefinedError as e:
            st.error(f"Erro de Variável Indefinida: {e}")
            st.info("Verifique se carregou as planilhas de contexto ou selecionou o auditado correto.")
        except Exception as e:
            st.error(f"Erro na Renderização: {e}")
    else:
        st.info("O editor está vazio.")
