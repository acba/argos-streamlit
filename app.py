import streamlit as st
import os

os.makedirs('tmp', exist_ok=True)

# Configuração da página principal
st.set_page_config(
    page_title="Argos - Auditoria Simplificada",
    layout="wide"
)

st.logo('img/argos_logo.png', size='large')


if 'configurado' not in st.session_state:
    st.session_state.configurado = True

# Inicializa o estado da sessão para manter os resultados após o download
# Isso garante que o estado seja compartilhado entre as páginas.
if 'files_processed' not in st.session_state:
    st.session_state.files_processed = False
if 'audit_completed' not in st.session_state:
    st.session_state.audit_completed = False
if 'audit_results' not in st.session_state:
    st.session_state.audit_results = None
if 'download_files' not in st.session_state:
    st.session_state.download_files = {}

home_page = st.Page("pages/home.py", title="Home", default=True)
aplica_procedimento_page = st.Page('pages/aplica_procedimentos.py', title="Aplica Procedimentos")
carrega_auditoria_page = st.Page('pages/carrega_auditoria.py', title="Carregar Resultado")
visualiza_resultados_page = st.Page('pages/visualiza_resultados.py', title="Visualiza Resultado")
gera_relatorios_individuais_page = st.Page('pages/gera_relatorios_individuais.py', title="Gera Relatórios Individuais")
gera_anexo_evidencias_page = st.Page('pages/gera_anexo_evidencias.py', title="Gera Anexo Evidências")

gera_questionario_comentarios_gestor_page = st.Page('pages/gera_questionario_comentarios_gestor.py', title="Gera Questionários do Comentários do Gestor")
gera_anexos_comentarios_gestor_page = st.Page('pages/gera_anexos_comentarios_gestor.py', title="Gera Anexos de Comentários do Gestor")

analise_gemini_auditados_page = st.Page('pages/analise_gemini.py', title="Análise de Auditados com IA")
analise_ia_geral_page = st.Page('pages/analise_ia_geral.py', title="Análise Geral com IA")

navigation_items = {
    "Procedimentos": [home_page, aplica_procedimento_page, carrega_auditoria_page],
    "Relatório": [],
    "Comentários do Gestor": [],
    "Análise IA": [],
}

# Adiciona páginas condicionalmente se a auditoria foi concluída
if st.session_state.audit_completed:
    navigation_items["Procedimentos"].append(visualiza_resultados_page)
    navigation_items["Relatório"].extend([gera_relatorios_individuais_page, gera_anexo_evidencias_page])
    navigation_items["Comentários do Gestor"].extend([gera_questionario_comentarios_gestor_page, gera_anexos_comentarios_gestor_page])
    navigation_items["Análise IA"].append(analise_gemini_auditados_page)

pg = st.navigation(navigation_items)

pg.run()
