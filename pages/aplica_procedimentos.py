import streamlit as st
import pandas as pd
import re
import os

from classes import FonteInformacao, AcaoVerificacao, ProcedimentoAuditoria, Auditado, \
    gerar_tabela_encaminhamentos, gerar_tabela_achados, gerar_tabela_situacoes_inconformes
from utils import carregar_dados

st.set_page_config(page_title="Aplicar Procedimentos", layout="wide")

st.title("Aplicar Procedimentos de Auditoria")
st.write("Esta seção realiza a aplicação dos procedimentos de auditoria definidos no mapa de verificação de achados. O processo gera como saída o objeto no modelo de dados do Audita, planilhas apresentando relação achados, encaminhamentos e situações encontrados por auditado, e relatórios individuais da aplicação dos procedimentos nos auditados.")

with st.expander("Modelo de Dados do Audita"):
    st.image('img/mermaid-diagram.png')
    # st.image('img/mermaid-diagram.svg')

# 1. Upload dos arquivos Excel
with st.expander("1. Carregar Arquivos Excel", expanded=True):
    col1, col2 = st.columns([3, 1])
    with col1:
        arquivo_auditados = st.file_uploader("Base de Auditados (ex: base-auditados.xlsx)", type=["xlsx"])
    with col2:
        st.write("")
        st.write("")
        with open("docs/bd_auditados.xlsx", "rb") as file:
            st.download_button(label="Baixar Exemplo", data=file, file_name="bd_auditados.xlsx", icon=":material/table:",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    col1, col2 = st.columns([3, 1])
    with col1:
        arquivo_mapa_achados = st.file_uploader("Mapa de Verificação e Achados (ex: mapa-verificacao-achados.xlsx)",
                                                type=["xlsx"])
    with col2:
        st.write("")
        st.write("")
        with open("docs/mapa-verificacao-achados.xlsx", "rb") as file:
            st.download_button(label="Baixar Exemplo", data=file, file_name="mapa-verificacao-achados.xlsx", icon=":material/table:",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    col1, col2 = st.columns([3, 1])
    with col1:
        arquivos_fontes_dados = st.file_uploader("Fontes de Informação (arquivos .xlsx)", type=["xlsx"],
                                                 accept_multiple_files=True)
    with col2:
        st.write("")
        st.write("")
        with open("docs/fonte_informacao.xlsx", "rb") as file:
            st.download_button(label="Baixar Exemplo", data=file, file_name="fonte_informacao.xlsx", icon=":material/table:",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if not arquivo_auditados or not arquivo_mapa_achados or not arquivos_fontes_dados:
        st.info("Por favor, carregue todos os arquivos Excel de entrada.")

# 2. Processamento dos dados
if arquivo_auditados and arquivo_mapa_achados and arquivos_fontes_dados:
    if st.button("Processar arquivos e gerar achados"):
        st.session_state.files_processed = False
        st.session_state.audit_completed = False
        st.session_state.audit_results = None

        try:
            if not st.session_state.files_processed:
                # Leitura dos DataFrames
                with st.spinner("Carregando planilhas..."):
                    df_jurisdicionados = carregar_dados(arquivo_auditados, skiprows=0)

                    df_procedimentos = carregar_dados(arquivo_mapa_achados, sheet_name='Procedimentos de Auditoria')
                    df_acoes_verificacao = carregar_dados(arquivo_mapa_achados, sheet_name='Ações de Verificação')
                    df_fontes = carregar_dados(arquivo_mapa_achados, sheet_name='Fontes de Informação')

                # Mapeia os arquivos de fonte de dados carregados pelo nome
                fontes_dados_carregadas = {f.name: f for f in arquivos_fontes_dados}

                # 1. Leitura das fontes de informação
                with st.spinner("Processando fontes de informação..."):
                    fontes = {}
                    for _, row in df_fontes.iterrows():
                        nome_arquivo_fonte = os.path.basename(row['filepath'])
                        fonte = FonteInformacao(
                            descricao=row['descricao'],
                            filepath=fontes_dados_carregadas.get(nome_arquivo_fonte),
                            chave_jurisdicionado=row['chave_jurisdicionado'],
                            id=row['id']
                        )
                        try:
                            fonte.read()
                            fontes[fonte.id] = fonte
                        except IOError as e:
                            st.error(f"Erro ao carregar a fonte de informação '{fonte.descricao}': {e}")
                        except AttributeError:
                            st.error(f"Arquivo da fonte de informação '{fonte.descricao}' ('{nome_arquivo_fonte}') não foi encontrado nos arquivos carregados. Verifique o nome.")

                # 2. Leitura das ações de verificação
                with st.spinner("Processando ações de verificação..."):
                    acoes = {}
                    for _, row in df_acoes_verificacao.iterrows():
                        fonte_informacao = fontes.get(row['id_fonte_informacao'])
                        if not fonte_informacao:
                            st.error(f"Ação de verificação '{row['id']}' refere-se a uma fonte de informação ('{row['id_fonte_informacao']}') que não foi encontrada.")
                            continue

                        acao = AcaoVerificacao(
                            fonte_informacao=fonte_informacao, informacao_requerida=row['informacao_requerida'],
                            acao_exclusiva_auditados=row['acao_exclusiva_auditados'], criterio=row['criterio'],
                            descricao_situacao_inconforme=row['descricao_situacao_inconforme'], descricao_evidencia=row['descricao_evidencia'],
                            situacao_inconforme=row['situacao_inconforme'], situacao_encontrada_nan_e_achado=row['situacao_encontrada_nan_e_achado'],
                            tipo_encaminhamento=row['tipo_encaminhamento'], encaminhamento=row['encaminhamento'],
                            pre_encaminhamento=row['pre_encaminhamento'], auditado_inexistente_e_achado=row['auditado_inexistente_e_achado'],
                            descricao_auditado_inexistente=row['descricao_auditado_inexistente'], id=row['id']
                        )
                        acoes[acao.id] = acao

                # 3. Leitura dos procedimentos de auditoria
                with st.spinner("Processando procedimentos de auditoria..."):
                    procedimentos = {}
                    for _, row in df_procedimentos.iterrows():
                        procedimento = ProcedimentoAuditoria(
                            descricao=row['descricao'], logica_achado=row['logica_achado'],
                            numero_achado=row['numero_achado'], nome_achado=row['nome_achado'], id=row['id']
                        )
                        acao_ids = [acao_id for acao_id in re.split(r'[\&\|\~\(\)\s]+', procedimento.logica_achado.replace("(", "").replace(")", "")) if acao_id]
                        for acao_id in acao_ids:
                            acao = acoes.get(acao_id.strip())
                            if acao:
                                procedimento.adicionar_acao(acao)
                        procedimentos[procedimento.id] = procedimento

                # 4. Leitura dos auditados
                with st.spinner("Carregando lista de auditados..."):
                    auditados = {}
                    for _, row in df_jurisdicionados.iterrows():
                        auditado = Auditado(nome=row['orgao'], sigla=row['sigla'])
                        auditados[auditado.sigla] = auditado

                st.success("Arquivos carregados e processados com sucesso!")
                st.session_state.files_processed = True

            # Se já carregou as planilhas mas ainda não finalizou a execução dos procedimentos
            if st.session_state.files_processed and not st.session_state.audit_completed:
                with st.spinner("Executando procedimentos de auditoria... Por favor, aguarde."):
                    # Execução da auditoria
                    for auditado in auditados.values():
                        auditado.aplicar_procedimentos(procedimentos.values(), debug=False)

                    # Geração das tabelas
                    tabela_encaminhamentos = gerar_tabela_encaminhamentos(auditados)
                    tabela_achados = gerar_tabela_achados(auditados)
                    tabela_situacoes = gerar_tabela_situacoes_inconformes(auditados)

                    st.session_state.audit_results = {
                        "auditados": auditados,
                        "tabela_encaminhamentos": tabela_encaminhamentos,
                        "tabela_achados": tabela_achados,
                        "tabela_situacoes": tabela_situacoes,
                    }
                    st.session_state.audit_completed = True
                    st.rerun()

        except ValueError as e:
            st.error(f"Erro de configuração: {e}")
        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante o processamento: {e}")

if st.session_state.audit_completed:
    st.success("Auditoria concluída! Navegue para 'Visualizar Resultado' ou 'Gerar Relatórios'.")