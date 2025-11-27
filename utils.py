import re
import streamlit as st
import pandas as pd
import jinja2
from jinja2 import Environment, BaseLoader, StrictUndefined
from google.genai import types
from docxtpl import InlineImage
from docx.shared import Mm
import logging
from datetime import date

def carregar_dados(filepath, sheet_name=0, skiprows=2):
    """Lê um arquivo Excel e retorna um DataFrame, tratando erros."""
    try:
        return pd.read_excel(filepath, sheet_name=sheet_name, skiprows=skiprows).map(lambda x: x.strip() if isinstance(x, str) else x)
    except Exception as e:
        st.error(f"Erro ao carregar a planilha '{sheet_name}': {e}")
        return None

def get_variaveis_template(template_md_content):
    """Coleta as variáveis presentes em um template Jinja2."""
    if not template_md_content:
        return set()
    env = Environment(loader=BaseLoader())
    ast = env.parse(template_md_content)
    return jinja2.meta.find_undeclared_variables(ast)

class StreamlitLogHandler(logging.Handler):
    """Handler de logging customizado para exibir logs do pypandoc no Streamlit."""
    def __init__(self, container):
        super().__init__()
        self.container = container
        self.records = []

    def emit(self, record):
        self.records.append(record)
        msg = self.format(record)
        self.container.warning(msg)


def parse_expression(expression):
    tokens = []
    current_token = ''

    for char in expression:
        if char in {'|', '&', '~', '(', ')'}:
            if current_token:
                tokens.append(current_token.strip())
                current_token = ''
            tokens.append(char)
        else:
            current_token += char

    if current_token:
        tokens.append(current_token.strip())

    return tokens

def infix_to_rpn(tokens):
    precedence = {'~': 3, '&': 2, '|': 1}
    output = []
    stack = []

    for token in tokens:
        if token == '(':
            # Abre parêntese
            stack.append(token)
        elif token == ')':
            # Fecha parêntese
            while stack and stack[-1] != '(':
                output.append(stack.pop())
            stack.pop()  # Remove o '('
        elif token in precedence:
            # Operador
            while stack and stack[-1] in precedence and precedence[stack[-1]] >= precedence[token]:
                output.append(stack.pop())
            stack.append(token)
        else:
            output.append(token)

    while stack:
        output.append(stack.pop())

    return list(filter(lambda x: x != '', output))

def avalia_expressao(expressao_achado, situacao_encontrada, debug=False):
    expressao_achado = str(expressao_achado)
    situacao_encontrada = str(situacao_encontrada)

    # Primeiro tenta se é o caso de um eval
    try:
        if debug:
            print(f"Testando se {situacao_encontrada} {expressao_achado} = {eval(f'{situacao_encontrada} {expressao_achado}')}")
        return eval(f'{situacao_encontrada} {expressao_achado}')
    except Exception as e:
        parsed_tokens = parse_expression(expressao_achado)
        tokens = infix_to_rpn(parsed_tokens)

        #print(tokens)

        pilha = []

        for token in tokens:
            if token == '|':
                # Operador OR
                y = pilha.pop()
                x = pilha.pop()
                pilha.append(x or y)
            elif token == '&':
                # Operador AND
                y = pilha.pop()
                x = pilha.pop()
                pilha.append(x and y)
            elif token == '~':
                x = pilha.pop()
                pilha.append(not x)
            else:
                # Aqui se remove o '.' no final da string pois não está padronizada as respostas.
                # Assim, encontra-se resposta terminando em '.' como 'Não adota' ou 'Não adota.'
                a = re.sub(r'\.$', '', token)
                b = re.sub(r'\.$', '', situacao_encontrada)
                if debug:
                    print(f"    Checa se {a} == {b} - ", a == b)
                pilha.append(a == b)

        if len(pilha) == 1:
            if debug:
                print('        Achado:', pilha[0])

            return pilha[0]
        else:
            raise ValueError("Expressão lógica inválida")

def processa_imagens_contexto(contexto, context_files_path_map, template_type, base_docx=None):
    """Substitui nomes de arquivos de imagem no contexto pelos caminhos ou objetos de imagem apropriados."""
    image_extensions = ('.png', '.jpg', '.jpeg', '.gif', '.bmp')

    # Itera sobre uma cópia dos itens para permitir a modificação do dicionário
    for key, value in list(contexto.items()):
        if isinstance(value, str) and value.lower().endswith(image_extensions):
            if value in context_files_path_map:
                image_path = context_files_path_map[value]
                if template_type == 'docx':
                    if base_docx is None:
                        st.error("O objeto base_docx é necessário para processar imagens em templates .docx")
                        continue
                    # Para docx, substitui pelo objeto InlineImage
                    contexto[key] = InlineImage(base_docx, image_path, width=Mm(160))
                elif template_type == 'md':
                    # Para markdown, substitui pelo caminho do arquivo
                    contexto[key] = image_path
            else:
                st.warning(f"Arquivo de imagem '{value}' para a variável '{key}' não encontrado. A imagem não será inserida.")
                contexto[key] = f"[Imagem '{value}' não encontrada]"

    return contexto

def cross_ref_figuras(template_str: str) -> str:
    """
    Processa um template de texto para numerar automaticamente as referências
    de figuras e modificar as legendas das imagens, preservando os
    atributos de formatação do Pandoc (ex: {width=10cm}).

    A função opera em duas passadas:

    1. Mapeamento:
       Encontra todas as declarações ({#fig:ID#}) e referências ([@fig:ID])
       na ordem em que aparecem para criar um mapa de numeração
       (ex: {'fig_01': 1, 'fig_02': 2}).

    2. Substituição:
       - Substitui referências de texto (ex: [@fig:fig_01] -> "Figura 1").
       - Encontra as linhas de imagem (ex: ![Texto]({{path}}){attrs}{#fig:ID#})
         e as substitui por "![Figura 1 - Texto]({{path}}){attrs}".
    """

    # --- Passa 1: Mapeamento ---

    figura_map = {}
    contador = 1

    # Regex combinada (NÃO MUDA)
    regex_combinado = r"(?:\{#fig:([^#]+)#\}|\[@fig:([^\]]+)\])"

    for match in re.finditer(regex_combinado, template_str):
        id_declaracao = match.group(1)
        id_referencia = match.group(2)
        fig_id = id_declaracao if id_declaracao else id_referencia

        if fig_id and fig_id not in figura_map:
            figura_map[fig_id] = contador
            contador += 1

    if not figura_map:
        return template_str

    # --- Passa 2: Substituições ---

    texto_processado = template_str

    # 1. Substituir referências de TEXTO (NÃO MUDA)
    regex_ref_texto = r"\[@fig:([^\]]+)\]"

    def substituir_ref_texto(match):
        fig_id = match.group(1)
        if fig_id in figura_map:
            return f"Figura {figura_map[fig_id]}"
        return match.group(0)

    texto_processado = re.sub(regex_ref_texto, substituir_ref_texto, texto_processado)

    # 2. Modificar linhas de IMAGEM e remover tags de declaração (MODIFICADO)

    # Regex ATUALIZADA:
    # Grupo 1: ![ (alt text) ]
    # Grupo 2: (path)
    # Grupo 3: (bloco de atributos opcional, ex: {width=10cm})
    # Grupo 4: {#fig: (ID) #}
    regex_imagem_decl = r"!\[([^\]]*)\](\([^)]*\))\s*(\{[^}]*\})?\s*\{#fig:([^#]+)#\}"

    def modificar_legenda_imagem(match):
        alt_text = match.group(1)
        path = match.group(2)
        attributes = match.group(3)  # O bloco {width=10cm}
        fig_id = match.group(4)

        if fig_id in figura_map:
            numero = figura_map[fig_id]

            # Se o grupo de atributos não for encontrado (None),
            # o transformamos em uma string vazia.
            attr_str = attributes if attributes else ""

            # Reconstrói a string: ![Figura X - Texto](path){atributos}
            return f"![Figura {numero} - {alt_text}]{path}{attr_str}"

        return match.group(0) # Failsafe

    texto_processado = re.sub(regex_imagem_decl, modificar_legenda_imagem, texto_processado)

    return texto_processado


def avalia_gemini(client, prompt_text: str, modelo, temperature, response_format_choice, file_objects = []):
    """
    Chama a API do Gemini com a configuração apropriada.
    Retorna a resposta do modelo e uma mensagem de erro (se houver).
    """
    try:
        contents = [prompt_text] + file_objects
        # st.info(f'Avaliando com o modelo {modelo}...') # Comentado para evitar chamadas Streamlit em utils

        generation_config = types.GenerateContentConfig(
            max_output_tokens=65536, # Usando o valor sugerido de 65536
            temperature=temperature,
        )

        if response_format_choice == 'Estruturada':
            generation_config.response_mime_type = 'application/json'

        # Cria o conteúdo para a API
        response = client.models.generate_content(
            model=modelo,
            contents=contents,
            config=generation_config
        )

        return response, None

    except Exception as e:
        return None, f"Erro ao chamar a API Gemini: {e}"

def data_hoje_abnt():
    hoje = date.today()

    # Lista de meses ABNT (Jan. Fev. Mar. Abr. Maio Jun. Jul. Ago. Set. Out. Nov. Dez.)
    # Note que 'maio' não tem ponto e é minúsculo na citação direta,
    # mas aqui usaremos minúsculo padrão.
    meses_abnt = {
        1: 'jan.', 2: 'fev.', 3: 'mar.', 4: 'abr.', 5: 'maio', 6: 'jun.',
        7: 'jul.', 8: 'ago.', 9: 'set.', 10: 'out.', 11: 'nov.', 12: 'dez.'
    }

    dia = hoje.day
    mes = meses_abnt[hoje.month]
    ano = hoje.year

    return f"{dia} {mes} {ano}"

def data_hoje():
    return date.today().strftime("%d/%m/%Y")