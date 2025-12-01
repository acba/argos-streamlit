import streamlit as st
import pandas as pd
import io
import zipfile
import logging
import pypandoc
import docx
import os
import tempfile
import re
import ast

from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm
from jinja2 import Environment, BaseLoader, StrictUndefined, exceptions

from classes import gerar_tabela_achados
from utils import get_variaveis_template, StreamlitLogHandler, processa_imagens_contexto, cross_ref_figuras, cross_ref_tabelas,\
data_hoje_abnt, data_hoje, aplicar_estilo_tabelas, processar_quebras_pagina, substituir_underline_pandoc

def consolidar_templates(base_content, all_files_content, processed_files=None):
    """
    Consolida templates substituindo tags {% include 'arquivo' %} pelo conteúdo do arquivo.
    """
    if processed_files is None:
        processed_files = set()

    # Regex para identificar tags de include: {% include 'arquivo.md' %} ou {% include "arquivo.md" %}
    # Captura o nome do arquivo dentro das aspas.
    pattern = re.compile(r"\{%-?\s*include\s+['\"](.+?)['\"]\s*-?%\}")

    def replace_match(match):
        filename = match.group(1)
        if filename in all_files_content:
            if filename in processed_files:
                return f"<!-- ERRO: Ciclo de inclusão detectado para '{filename}' -->"

            new_processed = processed_files.copy()
            new_processed.add(filename)

            return consolidar_templates(all_files_content[filename], all_files_content, new_processed)
        else:
            return f"<!-- ERRO: Arquivo '{filename}' não encontrado nos uploads -->"

    return pattern.sub(replace_match, base_content)

st.set_page_config(page_title="Gera Relatórios Individuais", layout="wide")

st.title("Gera Relatórios Individuais")
st.write("Esta seção permite a geração de relatórios personalizados a partir dos dados de auditoria processados, usando templates em formato Markdown/Jinja2.")

if st.session_state.audit_completed:
    results = st.session_state.audit_results
    auditados = results["auditados"]

    col1, col2 = st.columns(2)
    with col1:
        st.write("**Auditados**")
        df_auditados = pd.DataFrame([{'sigla': a.sigla, 'nome': a.nome} for a in auditados.values()]).set_index('sigla')
        st.dataframe(df_auditados, height=200)
    with col2:
        st.write("**Achados**")
        df_achados = pd.DataFrame([{'nome': nome} for nome in gerar_tabela_achados(auditados).columns])
        st.dataframe(df_achados, height=200)

    st.subheader("1. Forneça dados de contexto adicionais (Opcional)")

    with st.expander("Como funciona o Contexto Adicional?"):
        st.markdown("""
        Você pode fazer upload de planilhas Excel (`.xlsx`) para enriquecer o relatório com dados específicos de cada auditado.

        **Regras Básicas:**
        1. A planilha deve ter uma coluna chamada **`sigla`**. Esta coluna será usada para vincular os dados ao auditado correto.
        2. Cada outra coluna se tornará uma **variável** disponível no seu template Jinja2.
           - Ex: Uma coluna `gerente_responsavel` pode ser acessada no template como `{{ gerente_responsavel }}`.

        **Funcionalidade Avançada (Colunas com `*`):**
        Se o nome de uma coluna terminar com um asterisco (ex: `lista_pendencias*`), o sistema tentará interpretar o conteúdo das células dessa coluna como **estruturas de dados Python** (listas, dicionários, tuplas), em vez de simples textos.

        - **Exemplo na planilha:**
          - Coluna: `itens_revisados*`
          - Célula: `['Item A', 'Item B', 'Item C']`

        - **Uso no Template:**
          Como a variável agora é uma lista real, você pode iterar sobre ela:
          ```jinja
          Itens revisados:
          {% for item in itens_revisados_ %}
          - {{ item }}
          {% endfor %}
          ```
        *Nota: O asterisco é removido do nome da variável final.*
        """)

    arquivos_contexto = st.file_uploader("Carregar Planilhas de Contexto (.xlsx)", type=["xlsx"], accept_multiple_files=True, help="As planilhas devem ter uma coluna 'sigla' para identificar o auditado.")
    df_contexto_extra = None
    if arquivos_contexto:
        dfs_contexto = []
        all_columns = {}  # Dicionário para rastrear colunas: nome_coluna -> lista de arquivos

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
                                # Remove o asterisco do nome da coluna
                                df_temp.rename(columns={col: col.rstrip('*')}, inplace=True)
                            except Exception as e:
                                st.warning(f"Falha ao converter coluna especial '{col}' no arquivo '{arq.name}': {e}")

                    # Rastreia colunas para identificar duplicatas
                    for col in df_temp.columns:
                        if col not in all_columns:
                            all_columns[col] = []
                        all_columns[col].append(arq.name)

                    dfs_contexto.append(df_temp)
                else:
                    st.warning(f"Arquivo '{arq.name}' ignorado: Coluna 'sigla' não encontrada.")
            except Exception as e:
                st.error(f"Erro ao ler '{arq.name}': {e}")

        # Verifica e alerta sobre colunas duplicadas
        duplicates = {col: files for col, files in all_columns.items() if len(files) > 1}
        if duplicates:
            for col, files in duplicates.items():
                st.warning(f"A coluna '{col}' aparece em múltiplos arquivos: {', '.join(files)}. O valor do primeiro arquivo processado será mantido.")

        if dfs_contexto:
            # Concatena e agrupa por índice para unificar informações de diferentes arquivos
            # .first() prioriza os valores dos primeiros arquivos da lista em caso de conflito de colunas
            df_contexto_extra = pd.concat(dfs_contexto).groupby(level=0).first()
            st.write(f"**Contexto consolidado ({len(dfs_contexto)} arquivos):**")
            st.dataframe(df_contexto_extra.head())

    arquivos_fontes_contexto = st.file_uploader("Arquivos presentes na planilha de contexto", accept_multiple_files=True)

    st.subheader("2. Forneça o template do relatório")
    with st.expander("Informações sobre os dados do template"):
        st.markdown('''
        Para o template está disponível o objeto Auditado e seus objetos Achado.

        ### Auditado
        Representa a entidade que está sendo auditada. Este objeto gerencia a aplicação dos procedimentos de auditoria e armazena todos os resultados correspondentes.

        ##### Atributos
        - nome (str): O nome completo da entidade auditada.
        - sigla (str): A sigla ou nome completo da entidade, usada como chave principal na maioria das operações.
        - foi_auditado (bool): Uma flag que se torna True após a execução dos procedimentos de auditoria para esta entidade.
        - procedimentos_executados (list): Uma lista que armazena cópias dos objetos ProcedimentoAuditoria que foram executados para este auditado, contendo os resultados específicos (ações, situações encontradas, etc.).
        - tem_achados (bool): Uma flag que se torna True se qualquer um dos procedimentos executados resultar em um achado.

        #### Métodos

        - get_nomes_achados(): Retorna uma lista de strings formatadas ("1. Nome do Achado") com os nomes de todos os achados identificados para o auditado.
        - get_achados(): Retorna um dicionário onde as chaves são os identificadores dos achados (ex: "achado1") e os valores são os próprios objetos Achado.
        - get_achado_por_nome(nome_achado): Busca e retorna o objeto Achado correspondente a um nome de achado específico fornecido como string. Retorna None se não encontrar.
        - get_situacoes_inconformes(): Retorna uma lista consolidada com todas as strings de "situações inconformes" de todos os achados encontrados para o auditado.
        - get_encaminhamentos(): Retorna uma lista consolidada e sem duplicatas com todas as strings de "encaminhamentos" de todos os achados.
        - get_plano_acao(): Retorna uma lista de dicionários, ideal para gerar uma tabela de plano de ação. Cada dicionário contém o número do achado, o tipo de encaminhamento e a descrição do encaminhamento.

        ### Achado
        Representa um achado de auditoria específico, consolidando suas características, evidências e encaminhamentos.

        #### Atributos
        - numero (int ou str): O número identificador do achado.
        - nome (str): O nome ou título descritivo do achado.
        - situacoes_encontradas (list): Uma lista de strings, onde cada string descreve uma situação inconforme que contribuiu para a materialização do achado.
        - encaminhamentos (list): Uma lista de dicionários, onde cada dicionário representa um encaminhamento proposto (contendo chaves como tipo e encaminhamento).
        - evidencias (list): Uma lista de strings, onde cada string descreve uma evidência que suporta o achado.
        ''')

    with st.expander("Informações sobre como preencher o template"):
        st.markdown('''
        Tanto os templates `.md` quanto os `.docx` utilizam a linguagem de template **Jinja2**. A sintaxe é a mesma para ambos.

        ---

        #### 1. Exibindo Variáveis
        Para exibir o conteúdo de uma variável ou atributo de um objeto, utilize chaves duplas `{{ }}`.

        **Exemplo:**
        ```jinja
        ## Relatório para: {{ auditado.nome }} ({{ auditado.sigla }})

        O valor da variável de contexto `minha_variavel` é: {{ minha_variavel }}
        ```

        ---

        #### 2. Condicionais (If / Else)
        Para exibir blocos de texto apenas se uma condição for verdadeira, utilize `{% if %}`.

        **Exemplo:**
        ```jinja
        {% if auditado.tem_achados %}
        ### Achados Encontrados
        A auditoria encontrou achados para esta unidade.
        {% else %}
        **Nenhum achado foi encontrado para esta unidade.**
        {% endif %}
        ```

        ---

        #### 3. Laços de Repetição (For)
        Para iterar sobre uma lista (como a lista de achados ou evidências), utilize `{% for %}`.

        **Exemplo para listar todos os achados:**
        ```jinja
        {% for achado in auditado.get_achados().values() %}
        #### Achado {{ achado.numero }} - {{ achado.nome }}

        **Evidências:**
        {% for evidencia in achado.evidencias %}
        - {{ evidencia }}
        {% endfor %}

        {% endfor %}
        ```

        ---

        #### 4. Inclusão de Templates (apenas .md)
        Você pode dividir seu template em vários arquivos e incluí-los usando `{% include %}`.

        **Exemplo:**
        ```jinja
        {% include 'cabecalho.md' %}

        Conteúdo principal...

        {% include 'rodape.md' %}
        ```
        Certifique-se de fazer o upload de todos os arquivos referenciados.

        ---

        **Para templates `.docx` (docxtpl):** A sintaxe é idêntica. Você insere as tags `{{ ... }}` e `{% ... %}` diretamente no seu documento Word. Para criar uma tabela dinâmica, por exemplo, coloque a tag `{% for ... %}` na primeira célula de uma linha e a tag `{% endfor %}` na última célula da mesma linha. O `docxtpl` irá replicar a linha para cada item na sua lista.
        ''')

    col1, col2 = st.columns(2)
    with col1:
        arquivos_template_md = st.file_uploader("Carregue arquivos de template (.md, .jinja)", type=["md", "jinja"], accept_multiple_files=True)
    with col2:
        arquivo_template_docx = st.file_uploader("Carregue um arquivo de template (.docx)", type=["docx"])

    template_content = None
    template_type = None

    if arquivos_template_md:
        template_type = 'md'
        template_files = {}
        for arquivo in arquivos_template_md:
            try:
                template_files[arquivo.name] = arquivo.read().decode('utf-8')
            except Exception as e:
                st.error(f"Erro ao ler o arquivo {arquivo.name}: {e}")

        if template_files:
            base_filename = None
            if len(template_files) > 1:
                # Detecta automaticamente o arquivo base (aquele com mais 'includes')
                include_pattern = re.compile(r"\{%-?\s*include\s+['\"](.+?)['\"]\s*-?%\}")
                base_filename = max(template_files, key=lambda k: len(include_pattern.findall(template_files[k])))
                st.info(f"Arquivo base detectado automaticamente: **{base_filename}**")
            else:
                base_filename = list(template_files.keys())[0]

            if base_filename:
                try:
                    template_content = consolidar_templates(template_files[base_filename], template_files)
                except RecursionError:
                    st.error("Erro: Loop infinito de inclusão de templates detectado.")
                except Exception as e:
                    st.error(f"Erro na consolidação: {e}")

    # Se não for MD, verifica se é DOCX
    if not template_content and arquivo_template_docx:
        template_type = 'docx'
        try:
            doc = docx.Document(arquivo_template_docx)
            template_content = "\n".join([para.text for para in doc.paragraphs])
            template_content = template_content.replace("%p", "%").replace('‘', "'").replace('’', "'").replace('“', "'").replace('”', "'")
        except Exception as e:
            st.error(f"Erro ao ler o arquivo .docx: {e}")

    if template_content:
        with st.expander("Conteúdo do template consolidado:"):
            st.code(template_content)

        vars_template = get_variaveis_template(template_content)
        st.write("Variáveis encontradas no template:")
        st.code(f"{sorted(vars_template)}")

        st.subheader("3. Gere os relatórios")

        # Seleção de auditados
        opcoes_auditados = list(df_auditados.index)
        selecionados = st.multiselect(
            "Selecione os auditados para gerar o relatório (deixe vazio para selecionar todos):",
            options=opcoes_auditados,
            default=opcoes_auditados
        )

        if st.button("Gerar Relatórios Individuais"):
            siglas_selecionadas = selecionados if selecionados else opcoes_auditados

            if not siglas_selecionadas:
                st.warning("Nenhum auditado selecionado.")
            else:
                # Processamento do template markdown
                with st.spinner(f"Gerando {len(siglas_selecionadas)} relatórios individuais..."):
                    # --- Lógica para lidar com arquivos de contexto (incluindo ZIP) ---
                    env = Environment(loader=BaseLoader(), undefined=StrictUndefined)
                    template_ref_docx = 'docs/template-relatorio-individual.docx'
                    generation_log = st.expander("Log de Geração", expanded=True)
                    zip_buffer = io.BytesIO()

                    with tempfile.TemporaryDirectory() as tmp_dir:
                        # Prepara os arquivos de contexto UMA VEZ, fora do loop de auditados
                        context_files_path_map = {}
                        unzip_dir = os.path.join(tmp_dir, "unzipped_context")
                        os.makedirs(unzip_dir, exist_ok=True)

                        for uploaded_file in arquivos_fontes_contexto:
                            if uploaded_file.name.lower().endswith('.zip'):
                                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                                    zip_ref.extractall(unzip_dir)
                                # Mapeia os arquivos extraídos
                                for root, _, files in os.walk(unzip_dir):
                                    for filename in files:
                                        context_files_path_map[filename] = os.path.join(root, filename)
                            else:
                                # Salva arquivos individuais em um local temporário
                                temp_path = os.path.join(tmp_dir, uploaded_file.name)
                                with open(temp_path, "wb") as f:
                                    f.write(uploaded_file.getvalue())
                                context_files_path_map[uploaded_file.name] = temp_path

                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                            # Filtra o DataFrame apenas para os selecionados
                            df_filtrado = df_auditados.loc[siglas_selecionadas]

                            for sigla, row_auditado in df_filtrado.iterrows():

                                with generation_log:
                                    st.markdown(f"--- \n#### Processando: **{sigla}**")
                                    contexto = row_auditado.to_dict()
                                    contexto['sigla'] = sigla
                                    contexto['data_hoje_abnt'] = data_hoje_abnt()
                                    contexto['data_hoje'] = data_hoje()

                                    if df_contexto_extra is not None and sigla in df_contexto_extra.index:
                                        contexto.update(df_contexto_extra.loc[sigla].to_dict())
                                    contexto['auditado'] = auditados[sigla]

                                    vars_faltantes = set(vars_template) - set(contexto.keys())
                                    if len(vars_faltantes):
                                        st.warning(f'Atenção: As seguintes variáveis estão sendo utilizadas no template, mas não existem nos dados do contexto: {", ".join(vars_faltantes)}\nPara não impedir o processamento da geração, serão preenchidos dados vazios para essas variáveis.')
                                        for var in vars_faltantes:
                                            contexto[var] = []

                                    if template_type == 'md':
                                        try:
                                            # Processa as imagens para o contexto do Markdown
                                            contexto = processa_imagens_contexto(contexto, context_files_path_map, 'md')
                                            template_content_local = cross_ref_figuras(template_content)
                                            template_content_local = cross_ref_tabelas(template_content_local)
                                            template_content_local = processar_quebras_pagina(template_content_local)
                                            template_content_local = substituir_underline_pandoc(template_content_local)

                                            template_md = env.from_string(template_content_local)
                                            conteudo_final_md = template_md.render(contexto)

                                            md_filename = os.path.join(tmp_dir, f'_relatorio-{sigla}.md')
                                            with open(md_filename, 'w', encoding='utf-8') as f:
                                                f.write(conteudo_final_md)

                                            # Adiciona o MD processado ao ZIP para debug/conferência
                                            zip_f.write(md_filename, arcname=f'_Relatorio-{sigla}.md')

                                            docx_filename = os.path.join(tmp_dir, f'relatorio-{sigla}.docx')

                                            # Define caminhos de recursos para o Pandoc encontrar as imagens
                                            resource_paths = ['.', tmp_dir, unzip_dir]
                                            resource_path_arg = '--resource-path=' + os.pathsep.join(resource_paths)

                                            args_docx = ['--figure-caption-position=above', '--reference-doc=' + template_ref_docx, resource_path_arg]
                                            pypandoc_logger = logging.getLogger('pypandoc')
                                            pypandoc_logger.setLevel(logging.WARNING)
                                            log_container = st.empty()
                                            handler = StreamlitLogHandler(log_container)
                                            pypandoc_logger.addHandler(handler)

                                            pypandoc.convert_file(md_filename, to='docx', outputfile=docx_filename, extra_args=args_docx)
                                            pypandoc_logger.removeHandler(handler)

                                            # Aplica estilos nas tabelas
                                            aplicar_estilo_tabelas(docx_filename)

                                            if not handler.records: st.success(f"Relatório para **{sigla}** gerado.")
                                            zip_f.write(docx_filename, arcname=f'Relatorio-{sigla}.docx')

                                        except exceptions.UndefinedError as e:
                                            st.error(f"**Erro no template para `{sigla}`:** A variável `{e.message.split(' is undefined')[0]}` não foi encontrada.")
                                        except Exception as e:
                                            st.error(f"Erro ao gerar relatório para **{sigla}**: {e}")
                                    elif template_type == 'docx':
                                        try:
                                            base_docx = DocxTemplate(arquivo_template_docx)
                                            # Processa as imagens para o contexto do DOCX
                                            contexto = processa_imagens_contexto(contexto, context_files_path_map, 'docx', base_docx=base_docx)
                                            base_docx.render(contexto)
                                        except Exception as e:
                                            st.error(f"Erro ao renderizar o template DOCX para **{sigla}**: {e}")

                                        docx_filename = os.path.join(tmp_dir, f'relatorio-{sigla}.docx')
                                        base_docx.save(docx_filename)

                                        # Aplica estilos nas tabelas
                                        aplicar_estilo_tabelas(docx_filename)

                                        zip_f.write(docx_filename, arcname=f'Relatorio-{sigla}.docx')
                                        st.success(f"Relatório para **{sigla}** gerado.")

                                    else:
                                        st.error("Por favor, forneça um template (colando o texto ou carregando o arquivo).")

                    st.session_state.download_files['relatorios_individuais_zip'] = zip_buffer.getvalue()
                st.success("Geração de relatórios concluída!")
    if 'relatorios_individuais_zip' in st.session_state.download_files:
        st.download_button(
            label="Baixar Todos os Relatórios Individuais (.zip)",
            data=st.session_state.download_files['relatorios_individuais_zip'],
            file_name="relatorios_individuais.zip",
            mime="application/zip"
        )

else:
    st.info("Por favor, aplique os procedimentos ou carregue um resultado de auditoria antes de gerar relatórios.")