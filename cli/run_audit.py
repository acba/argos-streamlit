#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import re
import json
import logging
import argparse
import pandas as pd

# Ensure parent directory is in PYTHONPATH to import classes and utils
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from classes import FonteInformacao, AcaoVerificacao, ProcedimentoAuditoria, Auditado, \
    gerar_tabela_encaminhamentos, gerar_tabela_achados, gerar_tabela_situacoes_inconformes
from utils import aplicar_variaveis_temporarias, carregar_dados

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class NpEncoder(json.JSONEncoder):
    """JSON Encoder that converts numpy/pandas types to python primitives."""
    def default(self, obj):
        import numpy as np
        if isinstance(obj, np.integer):
            return int(obj)
        elif isinstance(obj, np.floating):
            return float(obj) if not pd.isna(obj) else None
        elif isinstance(obj, np.bool_):
            return bool(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        elif pd.isna(obj):
            return None
        return super().default(obj)

def main():
    parser = argparse.ArgumentParser(
        description='Executa os procedimentos de auditoria do Argos via Linha de Comando.'
    )
    parser.add_argument(
        '-a', '--auditados', required=True,
        help='Caminho para a base de auditados (.xlsx) (ex: bd_auditados.xlsx).'
    )
    parser.add_argument(
        '-m', '--mapa', required=True,
        help='Caminho para o mapa de verificação e achados (.xlsx) (ex: mapa-verificacao-achados.xlsx).'
    )
    parser.add_argument(
        '-f', '--fontes', nargs='+', required=True,
        help='Caminho para um ou mais arquivos de fontes de informação/respostas (.xlsx).'
    )
    parser.add_argument(
        '-j', '--output-json', default='.output_reports/resultado_auditoria.json',
        help='Caminho para salvar o resultado dos auditados em JSON (padrão: .output_reports/resultado_auditoria.json).'
    )
    parser.add_argument(
        '-x', '--output-xlsx', default='.output_reports/tabelas_consolidadas_auditoria.xlsx',
        help='Caminho para salvar as tabelas consolidadas em Excel (padrão: .output_reports/tabelas_consolidadas_auditoria.xlsx).'
    )

    args = parser.parse_args()

    # 1. Valida caminhos de entrada
    if not os.path.exists(args.auditados):
        logger.error(f"Arquivo de auditados '{args.auditados}' não encontrado.")
        sys.exit(1)
    if not os.path.exists(args.mapa):
        logger.error(f"Mapa de verificação '{args.mapa}' não encontrado.")
        sys.exit(1)

    # 2. Carrega planilhas do mapa
    logger.info("Carregando tabelas do mapa de verificação e auditados...")
    cols_jurisdicionados = ['sigla', 'orgao']
    cols_procedimentos = ['id', 'descricao', 'logica_achado', 'numero_achado', 'nome_achado']
    cols_acoes = ['id', 'id_fonte_informacao', 'informacao_requerida', 'criterio', 'situacao_inconforme', 'tipo_encaminhamento']
    cols_fontes = ['id', 'descricao', 'filepath', 'chave_jurisdicionado']
    cols_variaveis = ['id', 'id_fonte_informacao', 'nome', 'expressao', 'descricao']

    try:
        df_jurisdicionados = carregar_dados(args.auditados, skiprows=None, required_columns=cols_jurisdicionados)
        df_procedimentos = carregar_dados(args.mapa, sheet_name='Procedimentos de Auditoria', skiprows=None, required_columns=cols_procedimentos)
        df_acoes_verificacao = carregar_dados(args.mapa, sheet_name='Ações de Verificação', skiprows=None, required_columns=cols_acoes)
        df_fontes = carregar_dados(args.mapa, sheet_name='Fontes de Informação', skiprows=None, required_columns=cols_fontes)
        
        xl_mapa = pd.ExcelFile(args.mapa)
        if 'Variáveis Temporárias' in xl_mapa.sheet_names:
            df_variaveis_temporarias = carregar_dados(args.mapa, sheet_name='Variáveis Temporárias', skiprows=None, required_columns=cols_variaveis)
        else:
            df_variaveis_temporarias = pd.DataFrame(columns=cols_variaveis)
    except Exception as e:
        logger.error(f"Erro ao carregar arquivos de configuração da auditoria: {e}")
        sys.exit(1)

    # 3. Mapeia e carrega as fontes de informação
    logger.info("Carregando fontes de informação...")
    fontes_path_map = {os.path.basename(f): f for f in args.fontes}
    fontes = {}

    for _, row in df_fontes.iterrows():
        nome_arquivo_fonte = os.path.basename(row['filepath'])
        actual_path = fontes_path_map.get(nome_arquivo_fonte)
        
        if not actual_path:
            if len(args.fontes) == 1:
                actual_path = args.fontes[0]
                logger.warning(f"Mapeando fonte '{nome_arquivo_fonte}' para o único arquivo fornecido '{actual_path}'.")
            else:
                logger.error(f"Arquivo da fonte de informação '{nome_arquivo_fonte}' não foi encontrado nos arquivos fornecidos.")
                sys.exit(1)

        fonte = FonteInformacao(
            descricao=row['descricao'],
            filepath=actual_path,
            chave_jurisdicionado=row['chave_jurisdicionado'],
            id=row['id']
        )
        try:
            fonte.read()
            fontes[fonte.id] = fonte
        except Exception as e:
            logger.error(f"Erro ao ler a fonte de informação '{fonte.descricao}' ('{actual_path}'): {e}")
            sys.exit(1)

    # 4. Aplica variáveis temporárias
    logger.info("Aplicando variáveis temporárias...")
    try:
        aplicar_variaveis_temporarias(fontes, df_variaveis_temporarias)
    except Exception as e:
        logger.error(f"Erro ao aplicar variáveis temporárias: {e}")
        sys.exit(1)

    # 5. Inicializa as ações de verificação
    logger.info("Inicializando ações de verificação...")
    acoes = {}
    for _, row in df_acoes_verificacao.iterrows():
        fonte_informacao = fontes.get(row['id_fonte_informacao'])
        if not fonte_informacao:
            logger.error(f"Ação de verificação '{row['id']}' refere-se a uma fonte inexistente '{row['id_fonte_informacao']}'.")
            sys.exit(1)
        
        acao = AcaoVerificacao(
            fonte_informacao=fonte_informacao,
            informacao_requerida=row.get('informacao_requerida'),
            acao_exclusiva_auditados=row.get('acao_exclusiva_auditados'),
            criterio=row.get('criterio'),
            descricao_situacao_inconforme=row.get('descricao_situacao_inconforme'),
            descricao_evidencia=row.get('descricao_evidencia'),
            situacao_inconforme=row.get('situacao_inconforme'),
            situacao_encontrada_nan_e_achado=row.get('situacao_encontrada_nan_e_achado'),
            tipo_encaminhamento=row.get('tipo_encaminhamento'),
            encaminhamento=row.get('encaminhamento'),
            pre_encaminhamento=row.get('pre_encaminhamento'),
            auditado_inexistente_e_achado=row.get('auditado_inexistente_e_achado'),
            descricao_auditado_inexistente=row.get('descricao_auditado_inexistente'),
            id=row['id']
        )
        acoes[acao.id] = acao

    # 6. Inicializa os procedimentos de auditoria
    logger.info("Inicializando procedimentos de auditoria...")
    procedimentos = {}
    for _, row in df_procedimentos.iterrows():
        procedimento = ProcedimentoAuditoria(
            descricao=row['descricao'],
            logica_achado=row['logica_achado'],
            numero_achado=row['numero_achado'],
            nome_achado=row['nome_achado'],
            id=row['id']
        )
        # Parse logic to extract dependencies
        acao_ids = [acao_id for acao_id in re.split(r'[\&\|\~\(\)\s]+', procedimento.logica_achado.replace("(", "").replace(")", "")) if acao_id]
        for acao_id in acao_ids:
            acao = acoes.get(acao_id.strip())
            if acao:
                procedimento.adicionar_acao(acao)
        procedimentos[procedimento.id] = procedimento

    # 7. Inicializa os auditados
    logger.info("Carregando lista de auditados...")
    auditados = {}
    for _, row in df_jurisdicionados.iterrows():
        auditado = Auditado(nome=row['orgao'], sigla=row['sigla'])
        auditados[auditado.sigla] = auditado

    # 8. Executa a auditoria
    logger.info("Executando procedimentos de auditoria...")
    for auditado in auditados.values():
        auditado.aplicar_procedimentos(procedimentos.values(), debug=False)

    logger.info("Auditoria executada com sucesso! Gerando relatórios de saída...")

    # 9. Consolida tabelas
    tabela_encaminhamentos = gerar_tabela_encaminhamentos(auditados)
    tabela_achados = gerar_tabela_achados(auditados)
    tabela_situacoes = gerar_tabela_situacoes_inconformes(auditados)

    # Ranking de Auditados (como em visualiza_resultados)
    situations_per_auditado = {}
    achados_per_auditado = {}
    for sigla, auditado in auditados.items():
        if auditado.foi_auditado:
            situacoes = auditado.get_situacoes_inconformes()
            situations_per_auditado[sigla] = len(situacoes)
            achados_per_auditado[sigla] = len(auditado.get_nomes_achados())

    df_rank_combined = pd.DataFrame()
    if situations_per_auditado and achados_per_auditado:
        df_rank_situacoes = pd.DataFrame.from_dict(situations_per_auditado, orient='index', columns=['Qtd. Situações Inconformes'])
        df_rank_achados = pd.DataFrame.from_dict(achados_per_auditado, orient='index', columns=['Qtd. Achados Distintos'])
        df_rank_combined = df_rank_achados.join(df_rank_situacoes).sort_values(by=['Qtd. Achados Distintos', 'Qtd. Situações Inconformes'], ascending=False)

    # 10. Salva em JSON
    try:
        os.makedirs(os.path.dirname(os.path.abspath(args.output_json)), exist_ok=True)
        auditados_dict = {k: v.to_dict() for k, v in auditados.items()}
        with open(args.output_json, 'w', encoding='utf-8') as f:
            json.dump(auditados_dict, f, indent=2, ensure_ascii=False, cls=NpEncoder)
        logger.info(f"Contexto dos auditados salvo em: {args.output_json}")
    except Exception as e:
        logger.error(f"Erro ao salvar arquivo JSON de auditados: {e}")

    # 11. Salva em Excel
    try:
        os.makedirs(os.path.dirname(os.path.abspath(args.output_xlsx)), exist_ok=True)
        with pd.ExcelWriter(args.output_xlsx, engine='xlsxwriter') as writer:
            tabela_achados.to_excel(writer, sheet_name='Achados por Auditado')
            tabela_encaminhamentos.to_excel(writer, sheet_name='Encaminhamentos por Auditado')
            tabela_situacoes.to_excel(writer, sheet_name='Situações Inconformes')
            if not df_rank_combined.empty:
                df_rank_combined.to_excel(writer, sheet_name='Ranking de Auditados')
        logger.info(f"Tabelas consolidadas salvas em: {args.output_xlsx}")
    except Exception as e:
        logger.error(f"Erro ao salvar tabelas em Excel: {e}")

    logger.info("Auditoria finalizada!")

if __name__ == '__main__':
    main()
