import streamlit as st
import pandas as pd
import io
import json
import zipfile

st.set_page_config(page_title="Visualizar Resultado", layout="wide")

st.title("Visualizar Resultado de Auditoria")
st.write("Esta seção permite visualizar e baixar os resultados da auditoria processada.")

if st.session_state.audit_completed:
    results = st.session_state.audit_results

    st.header("Resultados Gerados")

    tab_graficos, tab_tabelas = st.tabs(["📊 Visualizar Gráficos", "📄 Visualizar Tabelas"])

    with tab_graficos:
        st.subheader("Quantitativo de Auditados por Achado")
        df_achados = results["tabela_achados"]
        achados_counts = (df_achados == 'X').sum().sort_values(ascending=True)
        st.bar_chart(achados_counts, horizontal=True)

        all_situations, all_encaminhamentos_texto, all_encaminhamentos_tipo = [], [], []
        situations_per_auditado, achados_per_auditado = {}, {}

        for sigla, auditado in results["auditados"].items():
            if auditado.foi_auditado:
                situacoes = auditado.get_situacoes_inconformes()
                all_situations.extend(situacoes)
                situations_per_auditado[sigla] = len(situacoes)
                achados_per_auditado[sigla] = len(auditado.get_nomes_achados())
                for p in auditado.procedimentos_executados:
                    if p.achado:
                        for e in p.achado.encaminhamentos:
                            all_encaminhamentos_texto.append(e['encaminhamento'].strip())
                            all_encaminhamentos_tipo.append(e['tipo'].strip())

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Distribuição por Tipo de Encaminhamento")
            if all_encaminhamentos_tipo:
                st.bar_chart(pd.Series(all_encaminhamentos_tipo).value_counts())
            else:
                st.info("Nenhum encaminhamento proposto.")

        with col2:
            st.subheader("Top 10 Encaminhamentos Mais Propostos")
            if all_encaminhamentos_texto:
                st.bar_chart(pd.Series(all_encaminhamentos_texto).value_counts().nlargest(10).sort_values(ascending=False), sort=False)
            else:
                st.info("Nenhum encaminhamento proposto.")

        st.subheader("Top 10 Situações Inconformes Mais Recorrentes")
        if all_situations:
            st.bar_chart(pd.Series(all_situations).value_counts().nlargest(10).sort_values(ascending=False), sort=False)
        else:
            st.info("Nenhuma situação inconforme encontrada.")

        st.subheader("Ranking de Auditados")
        if situations_per_auditado and achados_per_auditado:
            df_rank_situacoes = pd.DataFrame.from_dict(situations_per_auditado, orient='index', columns=['Qtd. Situações Inconformes'])
            df_rank_achados = pd.DataFrame.from_dict(achados_per_auditado, orient='index', columns=['Qtd. Achados Distintos'])
            df_rank_combined = df_rank_achados.join(df_rank_situacoes).sort_values(by=['Qtd. Achados Distintos', 'Qtd. Situações Inconformes'], ascending=False)
            st.dataframe(df_rank_combined)

    with tab_tabelas:
        st.subheader("Achados por Auditado")
        st.dataframe(results["tabela_achados"])
        st.subheader("Encaminhamentos por Auditado")
        st.dataframe(results["tabela_encaminhamentos"])
        st.subheader("Situações Inconformes por Auditado")
        st.dataframe(results["tabela_situacoes"])

    st.header("Baixar Resultados", divider="gray")

    if not st.session_state.download_files:
        with st.spinner("Gerando arquivos para download..."):
            # 1. Arquivo JSON (Substituindo Pickle)
            json_buffer = io.BytesIO()
            auditados_dict = {k: v.to_dict() for k, v in results["auditados"].items()}
            json_str = json.dumps(auditados_dict, indent=2, ensure_ascii=False)
            json_buffer.write(json_str.encode('utf-8'))
            st.session_state.download_files['json'] = json_buffer.getvalue()

            # 2. Arquivo Excel
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                results["tabela_achados"].to_excel(writer, sheet_name='Achados por Auditado')
                results["tabela_encaminhamentos"].to_excel(writer, sheet_name='Encaminhamentos por Auditado')
                results["tabela_situacoes"].to_excel(writer, sheet_name='Situações Inconformes')
            st.session_state.download_files['excel'] = excel_buffer.getvalue()

            # 3. Arquivos DOCX individuais e ZIP
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_f:
                st.session_state.download_files['docx'] = {}
                for sigla, auditado in results["auditados"].items():
                    if auditado.foi_auditado:
                        doc = auditado.documenta_procedimentos()
                        bio = io.BytesIO()
                        doc.save(bio)
                        docx_bytes = bio.getvalue()
                        st.session_state.download_files['docx'][sigla] = docx_bytes
                        zip_f.writestr(f"{auditado.sigla} - Relatorio.docx", docx_bytes)
            st.session_state.download_files['zip'] = zip_buffer.getvalue()

    st.download_button(
        label="Baixar Objeto Auditados (.json)",
        data=st.session_state.download_files['json'],
        file_name="auditados.json",
        mime="application/json"
    )
    st.download_button(
        label="Baixar Todas as Tabelas (.xlsx)",
        data=st.session_state.download_files['excel'],
        file_name="tabelas_consolidadas_auditoria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.download_button(
        label="Baixar Todos os Relatórios de Procedimentos (.zip)",
        data=st.session_state.download_files['zip'],
        file_name="relatorios_procedimentos_auditados.zip",
        mime="application/zip"
    )
    with st.expander("Baixar Relatórios de Procedimentos Individuais (.docx)"):
        for sigla, docx_bytes in st.session_state.download_files['docx'].items():
            st.download_button(label=f"Baixar Relatório {sigla}", data=docx_bytes, file_name=f"{sigla} - Relatorio.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"download_{sigla}")

else:
    st.info("Por favor, aplique os procedimentos ou carregue um resultado de auditoria antes de visualizar.")