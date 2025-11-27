import streamlit as st
import os
import zipfile
import io
from typing import List, Dict, Set
import pandas as pd
from classes import Auditado, Achado

class LimeSurveyGenerator:
    """Gera .lss consolidado espelhando EXATAMENTE o template DOCX"""

    def __init__(self):
        self.language = 'pt-BR'
        self.format = 'G'

    def _esc(self, text: str) -> str:
        """Escapa para CDATA"""
        if not text:
            return ''
        return str(text).replace(']]>', ']]]]><![CDATA[>')

    def generate_xml(self, auditados: List[Auditado], admin_email: str = 'brunosm@tcerj.tc.br') -> str:
        """Gera um único XML consolidado para todos os auditados"""
        if not auditados:
            return ""

        # Mapeia situações únicas e quais auditados as possuem
        # Chave: (numero_achado, nome_achado, texto_situacao)
        # Valor: Set de siglas
        situacoes_map: Dict[tuple, Set[str]] = {}
        # Para armazenar evidências acumuladas por situação (opcional, mas bom para contexto)
        evidencias_map: Dict[tuple, Set[str]] = {}

        for auditado in auditados:
            achados = [p.achado for p in auditado.procedimentos_executados if p.achado is not None]
            for achado in achados:
                sits = achado.situacoes_encontradas if achado.situacoes_encontradas else ['']

                for sit in sits:
                    key = (achado.numero, achado.nome, sit)

                    if key not in situacoes_map:
                        situacoes_map[key] = set()
                        evidencias_map[key] = set()

                    situacoes_map[key].add(auditado.sigla)

                    # Adiciona evidências do achado ao contexto da situação
                    for ev in achado.evidencias:
                        evidencias_map[key].add(ev)

        xml = self._build_xml(situacoes_map, evidencias_map, admin_email)
        return xml

    def _build_xml(self, situacoes_map: Dict[tuple, Set[str]], evidencias_map: Dict[tuple, Set[str]], admin_email: str) -> str:
        xml = '<?xml version="1.0" encoding="UTF-8"?>\n<document>\n'
        xml += ' <LimeSurveyDocType>Survey</LimeSurveyDocType>\n'
        xml += ' <DBVersion>363</DBVersion>\n'
        xml += ' <languages>\n  <language>pt-BR</language>\n </languages>\n'

        # Ordena as chaves para garantir ordem sequencial dos achados
        # Chave é (numero, nome, situacao). O numero pode ser string ou int, ideal garantir int para sort
        sorted_keys = sorted(situacoes_map.keys(), key=lambda x: (int(x[0]) if isinstance(x[0], int) or x[0].isdigit() else x[0], x[2]))

        # ORDEM CORRETA:
        xml += self._build_answers(sorted_keys)       # 1. answers
        xml += self._build_groups(sorted_keys, situacoes_map, evidencias_map)        # 2. groups
        xml += self._build_questions(sorted_keys)     # 3. questions
        xml += self._build_surveys(admin_email)              # 4. surveys
        xml += self._build_surveys_lang() # 5. surveys_languagesettings

        xml += '</document>'
        return xml

    def _build_groups(self, sorted_keys: List[tuple], situacoes_map: Dict[tuple, Set[str]], evidencias_map: Dict[tuple, Set[str]]) -> str:
        xml = ' <groups>\n  <fields>\n'
        xml += '   <fieldname>gid</fieldname>\n   <fieldname>sid</fieldname>\n'
        xml += '   <fieldname>group_name</fieldname>\n   <fieldname>group_order</fieldname>\n'
        xml += '   <fieldname>description</fieldname>\n   <fieldname>language</fieldname>\n'
        xml += '   <fieldname>randomization_group</fieldname>\n   <fieldname>grelevance</fieldname>\n'
        xml += '  </fields>\n  <rows>\n'

        # Grupo 1: Identificação (Sempre visível)
        xml += '   <row>\n'
        xml += '    <gid><![CDATA[1]]></gid>\n    <sid><![CDATA[0]]></sid>\n'
        xml += '    <group_name><![CDATA[Identificação]]></group_name>\n'
        xml += '    <group_order><![CDATA[1]]></group_order>\n'
        xml += '    <description><![CDATA[Por favor, preencha seus dados:]]></description>\n'
        xml += '    <language><![CDATA[pt-BR]]></language>\n'
        xml += '    <randomization_group/>\n    <grelevance/>\n'
        xml += '   </row>\n'

        gid = 2
        group_order = 2

        for key in sorted_keys:
            achado_num, achado_nome, situacao_texto = key
            siglas = sorted(list(situacoes_map[key]))
            evidencias = sorted(list(evidencias_map[key]))

            # Constrói a grelevance
            # TOKEN:FIRSTNAME == "SIGLA" OR ...
            conditions = [f'TOKEN:FIRSTNAME == "{sigla}"' for sigla in siglas]
            grelevance = " OR ".join(conditions)

            # Descrição completa
            desc = f'<h4>Achado {achado_num}: {self._esc(achado_nome)}</h4>'
            desc += '<p>O Relatório Individual apresenta e descreve as situações que ensejaram a ocorrência deste achado, as evidências levantadas, critérios aplicados as propostas de encaminhamento sugeridas pela Equipe de Auditoria.</p>'

            # Situação Específica deste grupo
            if situacao_texto:
                desc += f'<p><strong>Situação identificada:</strong> {self._esc(situacao_texto)}</p>'

            # # Evidências
            # if evidencias:
            #     desc += '<p><strong>Evidências do Achado:</strong></p><ul>'
            #     for ev in evidencias:
            #         desc += f'<li>{self._esc(ev)}</li>'
            #     desc += '</ul>'

            # Instruções
            desc += '<hr/>'
            desc += '<p><strong>INSTRUÇÕES:</strong></p>'
            desc += '<ul>'
            desc += '<li>Se <strong>concordar</strong>: Use o campo "Comentários" para informar ações corretivas, avaliação sobre as propostas e consequências práticas.</li>'
            desc += '<li>Se <strong>discordar</strong>: Use o campo "Justificativa" para apresentar as razões da discordância e envie documentos comprobatórios.</li>'
            desc += '</ul>'

            group_name = f'Achado {achado_num}'

            xml += '   <row>\n'
            xml += f'    <gid><![CDATA[{gid}]]></gid>\n    <sid><![CDATA[0]]></sid>\n'
            xml += f'    <group_name><![CDATA[{group_name}]]></group_name>\n'
            xml += f'    <group_order><![CDATA[{group_order}]]></group_order>\n'
            xml += f'    <description><![CDATA[{desc}]]></description>\n'
            xml += '    <language><![CDATA[pt-BR]]></language>\n'
            xml += '    <randomization_group/>\n'
            xml += f'    <grelevance><![CDATA[{grelevance}]]></grelevance>\n'
            xml += '   </row>\n'

            gid += 1
            group_order += 1

        xml += '  </rows>\n </groups>\n'
        return xml

    def _build_questions(self, sorted_keys: List[tuple]) -> str:
        xml = ' <questions>\n  <fields>\n'
        xml += '   <fieldname>qid</fieldname>\n   <fieldname>parent_qid</fieldname>\n'
        xml += '   <fieldname>sid</fieldname>\n   <fieldname>gid</fieldname>\n'
        xml += '   <fieldname>type</fieldname>\n   <fieldname>title</fieldname>\n'
        xml += '   <fieldname>question</fieldname>\n   <fieldname>preg</fieldname>\n'
        xml += '   <fieldname>help</fieldname>\n   <fieldname>other</fieldname>\n'
        xml += '   <fieldname>mandatory</fieldname>\n   <fieldname>question_order</fieldname>\n'
        xml += '   <fieldname>language</fieldname>\n   <fieldname>scale_id</fieldname>\n'
        xml += '   <fieldname>same_default</fieldname>\n   <fieldname>relevance</fieldname>\n'
        xml += '   <fieldname>modulename</fieldname>\n'
        xml += '  </fields>\n  <rows>\n'

        qid = 1

        # Identificação (4 perguntas) - Grupo 1
        for title, question, mandatory in [
            ('nome', 'Nome completo', 'Y'),
            ('cargo', 'Cargo/Função', 'Y'),
            ('email', 'E-mail', 'Y'),
            ('telefone', 'Telefone (opcional)', 'N')
        ]:
            xml += self._q_row(qid, 1, title, question, 'S', mandatory, qid)
            qid += 1

        # Achados
        gid = 2
        # Contador global para gerar códigos únicos de pergunta se necessário,
        # ou usamos a tupla (Achado, Situação). Como sorted_keys é linear, podemos usar index.

        for idx, key in enumerate(sorted_keys, 1):
            achado_num, achado_nome, situacao_texto = key
            q_order = 1

            # Define base code for this situation.
            # Usar um hash ou contador sequencial para garantir unicidade e brevidade.
            # Ex: A{Num}G{Index}
            base_code = f'A{achado_num}G{idx}'

            # P1: Concordância (4 opções)
            if situacao_texto:
                q_text = f'No que tange à situação "<strong>{self._esc(situacao_texto)}</strong>" apontada no Relatório Individual, a organização:'
            else:
                q_text = f'Em relação ao Achado {achado_num}, a organização:'

            xml += self._q_row(qid, gid, f'{base_code}Conc', q_text, 'L', 'Y', q_order)
            qid += 1
            q_order += 1

            # Logic expressions
            relevance_conc = f'(({base_code}Conc.NAOK == "SQ001" or {base_code}Conc.NAOK == "SQ002" or {base_code}Conc.NAOK == "SQ003"))'
            relevance_disc = f'(({base_code}Conc.NAOK == "SQ006"))'
            relevance_evi  = f'(({base_code}Conc.NAOK == "SQ001" or {base_code}Conc.NAOK == "SQ002" or {base_code}Conc.NAOK == "SQ006"))'

            # P2: Comentários (se concordar)
            q_text_coment = '(Opcional) Caso desejável, encaminhem seus comentários contemplando a perspectiva da organização e as ações corretivas que pretendem tomar, bem como avaliação sobre as propostas de determinação e/ou recomendação formuladas, informando sobre as consequências práticas de sua implementação e eventuais alternativas.'
            xml += self._q_row(qid, gid, f'{base_code}Com',
                q_text_coment,
                'T', 'N', q_order,
                'Se concordar de alguma forma com o achado, use este campo para tecer esclarecimentos.',
                relevance=relevance_conc)
            qid += 1
            q_order += 1

            # P3: Justificativa (se discordar)
            q_text_justif = 'Ajude-nos a entender por que você escolheu a opção acima'
            xml += self._q_row(qid, gid, f'{base_code}Jus',
                q_text_justif,
                'T', 'Y', q_order,
                'Se discordar, apresente as razões da sua discordância.',
                relevance=relevance_disc)
            qid += 1
            q_order += 1

            # P4: Documentos
            xml += self._q_row(qid, gid, f'{base_code}Evi',
                'Caso necessário, envie documentos que evidenciem suas justificativas',
                '|', 'N', q_order,
                'É aceito arquivo com extensão PDF ou ZIP. Caso haja mais de um arquivo, compactar em formato ZIP.',
                relevance=relevance_evi)
            qid += 1
            q_order += 1

            # Avança Grupo
            gid += 1

        xml += '  </rows>\n </questions>\n'
        return xml

    def _q_row(self, qid: int, gid: int, title: str, question: str, qtype: str,
               mandatory: str, order: int, help_text: str = '', relevance: str = '1') -> str:
        """Gera linha de pergunta"""
        xml = '   <row>\n'
        xml += f'    <qid><![CDATA[{qid}]]></qid>\n'
        xml += '    <parent_qid><![CDATA[0]]></parent_qid>\n'
        xml += '    <sid><![CDATA[0]]></sid>\n'
        xml += f'    <gid><![CDATA[{gid}]]></gid>\n'
        xml += f'    <type><![CDATA[{qtype}]]></type>\n'
        xml += f'    <title><![CDATA[{title}]]></title>\n'
        xml += f'    <question><![CDATA[{question}]]></question>\n'
        xml += '    <preg/>\n'
        xml += f'    <help><![CDATA[{help_text}]]></help>\n'
        xml += '    <other><![CDATA[N]]></other>\n'
        xml += f'    <mandatory><![CDATA[{mandatory}]]></mandatory>\n'
        xml += f'    <question_order><![CDATA[{order}]]></question_order>\n'
        xml += '    <language><![CDATA[pt-BR]]></language>\n'
        xml += '    <scale_id><![CDATA[0]]></scale_id>\n'
        xml += '    <same_default><![CDATA[0]]></same_default>\n'
        xml += f'    <relevance><![CDATA[{relevance}]]></relevance>\n'
        xml += '    <modulename/>\n'
        xml += '   </row>\n'
        return xml

    def _build_answers(self, sorted_keys: List[tuple]) -> str:
        """Gera opções de resposta (4 opções de concordância)"""
        xml = ' <answers>\n  <fields>\n'
        xml += '   <fieldname>qid</fieldname>\n   <fieldname>code</fieldname>\n'
        xml += '   <fieldname>answer</fieldname>\n   <fieldname>sortorder</fieldname>\n'
        xml += '   <fieldname>assessment_value</fieldname>\n   <fieldname>language</fieldname>\n'
        xml += '   <fieldname>scale_id</fieldname>\n'
        xml += '  </fields>\n  <rows>\n'

        # QID inicial: 5 (após as 4 perguntas de identificação)
        qid = 5

        for _ in sorted_keys:
            # Para a primeira pergunta de cada situação (concordância)
            for code, answer, sort in [
                ('SQ001', 'Concorda e já atendeu às propostas de encaminhamento', 1),
                ('SQ002', 'Concorda e já está atendendo às propostas de encaminhamento', 2),
                ('SQ003', 'Concorda, mas ainda não adotou nenhuma medida para atender às propostas de encaminhamento', 3),
                ('SQ006', 'Discorda da sinalização de inadequação', 4)
            ]:
                xml += '   <row>\n'
                xml += f'    <qid><![CDATA[{qid}]]></qid>\n'
                xml += f'    <code><![CDATA[{code}]]></code>\n'
                xml += f'    <answer><![CDATA[{answer}]]></answer>\n'
                xml += f'    <sortorder><![CDATA[{sort}]]></sortorder>\n'
                xml += '    <assessment_value><![CDATA[0]]></assessment_value>\n'
                xml += '    <language><![CDATA[pt-BR]]></language>\n'
                xml += '    <scale_id><![CDATA[0]]></scale_id>\n'
                xml += '   </row>\n'

            # Avança 4 perguntas por situação
            qid += 4

        xml += '  </rows>\n </answers>\n'
        return xml

    def _build_surveys(self, admin_email: str) -> str:
        xml = ' <surveys>\n  <fields>\n'
        xml += '   <fieldname>sid</fieldname>\n   <fieldname>gsid</fieldname>\n'
        xml += '   <fieldname>admin</fieldname>\n   <fieldname>expires</fieldname>\n'
        xml += '   <fieldname>startdate</fieldname>\n   <fieldname>adminemail</fieldname>\n'
        xml += '   <fieldname>anonymized</fieldname>\n   <fieldname>faxto</fieldname>\n'
        xml += '   <fieldname>format</fieldname>\n   <fieldname>savetimings</fieldname>\n'
        xml += '   <fieldname>template</fieldname>\n   <fieldname>language</fieldname>\n'
        xml += '   <fieldname>additional_languages</fieldname>\n   <fieldname>datestamp</fieldname>\n'
        xml += '   <fieldname>usecookie</fieldname>\n   <fieldname>allowregister</fieldname>\n'
        xml += '   <fieldname>allowsave</fieldname>\n   <fieldname>autonumber_start</fieldname>\n'
        xml += '   <fieldname>autoredirect</fieldname>\n   <fieldname>allowprev</fieldname>\n'
        xml += '   <fieldname>printanswers</fieldname>\n   <fieldname>ipaddr</fieldname>\n'
        xml += '   <fieldname>refurl</fieldname>\n   <fieldname>showsurveypolicynotice</fieldname>\n'
        xml += '   <fieldname>publicstatistics</fieldname>\n   <fieldname>publicgraphs</fieldname>\n'
        xml += '   <fieldname>listpublic</fieldname>\n   <fieldname>htmlemail</fieldname>\n'
        xml += '   <fieldname>sendconfirmation</fieldname>\n   <fieldname>tokenanswerspersistence</fieldname>\n'
        xml += '   <fieldname>assessments</fieldname>\n   <fieldname>usecaptcha</fieldname>\n'
        xml += '   <fieldname>usetokens</fieldname>\n   <fieldname>bounce_email</fieldname>\n'
        xml += '   <fieldname>emailresponseto</fieldname>\n   <fieldname>emailnotificationto</fieldname>\n'
        xml += '   <fieldname>tokenlength</fieldname>\n   <fieldname>showxquestions</fieldname>\n'
        xml += '   <fieldname>showgroupinfo</fieldname>\n   <fieldname>shownoanswer</fieldname>\n'
        xml += '   <fieldname>showqnumcode</fieldname>\n   <fieldname>bounceprocessing</fieldname>\n'
        xml += '   <fieldname>showwelcome</fieldname>\n   <fieldname>showprogress</fieldname>\n'
        xml += '   <fieldname>questionindex</fieldname>\n   <fieldname>navigationdelay</fieldname>\n'
        xml += '   <fieldname>nokeyboard</fieldname>\n   <fieldname>alloweditaftercompletion</fieldname>\n'
        xml += '  </fields>\n  <rows>\n   <row>\n'
        xml += '    <sid><![CDATA[0]]></sid>\n    <gsid><![CDATA[1]]></gsid>\n'
        xml += '    <admin><![CDATA[TCE-RJ]]></admin>\n    <expires/>\n    <startdate/>\n'
        xml += f'    <adminemail><![CDATA[{admin_email}]]></adminemail>\n'
        xml += '    <anonymized><![CDATA[N]]></anonymized>\n    <faxto/>\n'
        xml += '    <format><![CDATA[G]]></format>\n'
        xml += '    <savetimings><![CDATA[N]]></savetimings>\n'
        xml += '    <template><![CDATA[fruity]]></template>\n'
        xml += '    <language><![CDATA[pt-BR]]></language>\n    <additional_languages/>\n'
        xml += '    <datestamp><![CDATA[Y]]></datestamp>\n'
        xml += '    <usecookie><![CDATA[N]]></usecookie>\n'
        xml += '    <allowregister><![CDATA[N]]></allowregister>\n'
        xml += '    <allowsave><![CDATA[Y]]></allowsave>\n'
        xml += '    <autonumber_start><![CDATA[0]]></autonumber_start>\n'
        xml += '    <autoredirect><![CDATA[N]]></autoredirect>\n'
        xml += '    <allowprev><![CDATA[Y]]></allowprev>\n'
        xml += '    <printanswers><![CDATA[Y]]></printanswers>\n'
        xml += '    <ipaddr><![CDATA[N]]></ipaddr>\n    <refurl><![CDATA[N]]></refurl>\n'
        xml += '    <showsurveypolicynotice><![CDATA[0]]></showsurveypolicynotice>\n'
        xml += '    <publicstatistics><![CDATA[N]]></publicstatistics>\n'
        xml += '    <publicgraphs><![CDATA[N]]></publicgraphs>\n'
        xml += '    <listpublic><![CDATA[N]]></listpublic>\n'
        xml += '    <htmlemail><![CDATA[Y]]></htmlemail>\n'
        xml += '    <sendconfirmation><![CDATA[Y]]></sendconfirmation>\n'
        xml += '    <tokenanswerspersistence><![CDATA[Y]]></tokenanswerspersistence>\n'
        xml += '    <assessments><![CDATA[N]]></assessments>\n'
        xml += '    <usecaptcha><![CDATA[N]]></usecaptcha>\n'
        xml += '    <usetokens><![CDATA[N]]></usetokens>\n'
        xml += f'    <bounce_email><![CDATA[{admin_email}]]></bounce_email>\n'
        xml += '    <emailresponseto/>\n    <emailnotificationto/>\n'
        xml += '    <tokenlength><![CDATA[15]]></tokenlength>\n'
        xml += '    <showxquestions><![CDATA[N]]></showxquestions>\n'
        xml += '    <showgroupinfo><![CDATA[B]]></showgroupinfo>\n'
        xml += '    <shownoanswer><![CDATA[N]]></shownoanswer>\n'
        xml += '    <showqnumcode><![CDATA[X]]></showqnumcode>\n'
        xml += '    <bounceprocessing><![CDATA[N]]></bounceprocessing>\n'
        xml += '    <showwelcome><![CDATA[Y]]></showwelcome>\n'
        xml += '    <showprogress><![CDATA[Y]]></showprogress>\n'
        xml += '    <questionindex><![CDATA[0]]></questionindex>\n'
        xml += '    <navigationdelay><![CDATA[0]]></navigationdelay>\n'
        xml += '    <nokeyboard><![CDATA[N]]></nokeyboard>\n'
        xml += '    <alloweditaftercompletion><![CDATA[N]]></alloweditaftercompletion>\n'
        xml += '   </row>\n  </rows>\n </surveys>\n'
        return xml

    def _build_surveys_lang(self) -> str:
        welcome = f'''<p>Prezado(a) Gestor(a) da {{TOKEN:FIRSTNAME}},</p>
<p>Este questionário é o instrumento pelo qual a {{TOKEN:FIRSTNAME}} pode se manifestar acerca dos achados apresentados pela Equipe de Auditoria no Relatório Individual Preliminar.</p>
<p>Para os itens que houver discordância em relação à situação encontrada, é necessário justificar o motivo da divergência no campo apropriado, além de encaminhar documentação que corrobore com a argumentação apresentada.</p>
<p><strong>Suas respostas são fundamentais para o processo de auditoria.</strong></p>'''

        xml = ' <surveys_languagesettings>\n  <fields>\n'
        xml += '   <fieldname>surveyls_survey_id</fieldname>\n   <fieldname>surveyls_language</fieldname>\n'
        xml += '   <fieldname>surveyls_title</fieldname>\n   <fieldname>surveyls_description</fieldname>\n'
        xml += '   <fieldname>surveyls_welcometext</fieldname>\n   <fieldname>surveyls_endtext</fieldname>\n'
        xml += '   <fieldname>surveyls_url</fieldname>\n   <fieldname>surveyls_urldescription</fieldname>\n'
        xml += '   <fieldname>surveyls_email_invite_subj</fieldname>\n   <fieldname>surveyls_email_invite</fieldname>\n'
        xml += '   <fieldname>surveyls_email_remind_subj</fieldname>\n   <fieldname>surveyls_email_remind</fieldname>\n'
        xml += '   <fieldname>surveyls_email_register_subj</fieldname>\n   <fieldname>surveyls_email_register</fieldname>\n'
        xml += '   <fieldname>surveyls_email_confirm_subj</fieldname>\n   <fieldname>surveyls_email_confirm</fieldname>\n'
        xml += '   <fieldname>surveyls_dateformat</fieldname>\n   <fieldname>surveyls_numberformat</fieldname>\n'
        xml += '   <fieldname>email_admin_notification_subj</fieldname>\n   <fieldname>email_admin_notification</fieldname>\n'
        xml += '   <fieldname>email_admin_responses_subj</fieldname>\n   <fieldname>email_admin_responses</fieldname>\n'
        xml += '  </fields>\n  <rows>\n   <row>\n'
        xml += '    <surveyls_survey_id><![CDATA[0]]></surveyls_survey_id>\n'
        xml += '    <surveyls_language><![CDATA[pt-BR]]></surveyls_language>\n'
        xml += f'    <surveyls_title><![CDATA[Comentários do Gestor - Auditoria]]></surveyls_title>\n'
        xml += f'    <surveyls_description><![CDATA[Auditoria]]></surveyls_description>\n'
        xml += f'    <surveyls_welcometext><![CDATA[{welcome}]]></surveyls_welcometext>\n'
        xml += '    <surveyls_endtext><![CDATA[<p><strong>Agradecemos sua participação!</strong></p>]]></surveyls_endtext>\n'
        xml += '    <surveyls_url/>\n    <surveyls_urldescription/>\n'
        xml += f'    <surveyls_email_invite_subj><![CDATA[Questionário - Auditoria]]></surveyls_email_invite_subj>\n'
        xml += '    <surveyls_email_invite/>\n    <surveyls_email_remind_subj/>\n'
        xml += '    <surveyls_email_remind/>\n    <surveyls_email_register_subj/>\n'
        xml += '    <surveyls_email_register/>\n    <surveyls_email_confirm_subj/>\n'
        xml += '    <surveyls_email_confirm/>\n'
        xml += '    <surveyls_dateformat><![CDATA[5]]></surveyls_dateformat>\n'
        xml += '    <surveyls_numberformat><![CDATA[0]]></surveyls_numberformat>\n'
        xml += '    <email_admin_notification_subj/>\n    <email_admin_notification/>\n'
        xml += '    <email_admin_responses_subj/>\n    <email_admin_responses/>\n'
        xml += '   </row>\n  </rows>\n </surveys_languagesettings>\n'
        return xml

st.set_page_config(page_title="Gerar Comentários do Gestor", layout="wide")

st.title("Gerar Questionário de Comentários do Gestor")
st.write("Esta ferramenta gera um arquivo .lss (LimeSurvey) consolidado, onde as perguntas são exibidas condicionalmente aos gestores (identificados pelo FIRSTNAME) conforme os achados da auditoria.")

if "audit_completed" in st.session_state and st.session_state.audit_completed:
    results = st.session_state.audit_results
    auditados = results["auditados"]

    # Filtra apenas auditados com achados
    auditados_com_achados = {k: v for k, v in auditados.items() if v.tem_achados}

    st.write(f"Total de auditados processados: **{len(auditados)}**")
    st.write(f"Auditados com achados (elegíveis para questionário): **{len(auditados_com_achados)}**")

    if not auditados_com_achados:
        st.warning("Nenhum auditado possui achados para gerar questionário.")
    else:
        # Seleção de auditados
        opcoes_auditados = list(auditados_com_achados.keys())
        selecionados = st.multiselect(
            "Selecione os auditados para incluir no questionário:",
            options=opcoes_auditados,
            default=opcoes_auditados
        )

        admin_email = st.text_input("Email do Administrador:", value="auditoria_seginfo@tcerj.tc.br")

        if st.button("Gerar Arquivo Consolidado .lss"):
            generator = LimeSurveyGenerator()

            # Se nenhum selecionado, usa todos
            siglas_selecionadas = selecionados if selecionados else opcoes_auditados
            lista_auditados = [auditados_com_achados[sigla] for sigla in siglas_selecionadas]

            with st.spinner("Gerando arquivo consolidado..."):
                try:
                    xml_content = generator.generate_xml(lista_auditados, admin_email)

                    if xml_content:
                        filename = "questionario_unificado.lss"
                        st.success(f"Arquivo consolidado gerado com sucesso! Contém lógica para {len(lista_auditados)} auditados.")

                        st.download_button(
                            label="Baixar Questionário Unificado (.lss)",
                            data=xml_content,
                            file_name=filename,
                            mime="application/xml"
                        )
                    else:
                        st.warning("O conteúdo gerado está vazio.")

                except Exception as e:
                    st.error(f"Erro ao gerar arquivo: {e}")
                    # st.exception(e) # Para debug

else:
    st.info("Por favor, carregue ou processe a auditoria primeiro na página 'Carregar Auditoria' ou 'Aplica Procedimentos'.")

