import streamlit as st
from utils import apply_custom_style, create_card

st.set_page_config(page_title="Argos - Home", layout="wide", page_icon="🛡️")

# Aplica o estilo global
apply_custom_style()

# Hero Section
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("""
    # Bem-vindo ao Argos
    ### Auditoria Simplificada e Inteligente

    O Argos é uma plataforma projetada para automatizar e aprimorar processos de auditoria.
    Com integração de IA e ferramentas de análise de dados, ele transforma a complexidade
    da conformidade em insights claros e acionáveis.
    """, unsafe_allow_html=True)

    st.write("")
    if st.button("Iniciar Nova Auditoria", type="primary"):
        st.switch_page("pages/aplica_procedimentos.py")

with col2:
    st.image("img/argos_logo.png", width=250)

st.markdown("---")

# Features Section
st.subheader("O que você pode fazer?")

row1_col1, row1_col2, row1_col3 = st.columns(3)

with row1_col1:
    st.markdown(create_card(
        "Aplicar Procedimentos",
        "Carregue seus dados e mapas de verificação para executar automaticamente procedimentos de auditoria complexos.",
        "⚙️"
    ), unsafe_allow_html=True)

with row1_col2:
    st.markdown(create_card(
        "Análise com IA",
        "Utilize o Google Gemini para analisar contextos, documentos e obter insights qualitativos sobre os auditados.",
        "🧠"
    ), unsafe_allow_html=True)

with row1_col3:
    st.markdown(create_card(
        "Gerar Relatórios",
        "Exporte relatórios detalhados, matrizes de achados e anexos de evidências prontos para uso.",
        "📄"
    ), unsafe_allow_html=True)

st.write("")
st.write("")

# Quick Links / Status
st.subheader("Status do Sistema")
if 'audit_completed' in st.session_state and st.session_state.audit_completed:
    st.success("✅ Uma auditoria está carregada na memória e pronta para análise.")
else:
    st.info("ℹ️ Nenhuma auditoria ativa no momento. Inicie aplicando procedimentos ou carregando um resultado salvo.")