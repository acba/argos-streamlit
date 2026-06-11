# Argos - Interface de Linha de Comando (CLI)

Esta pasta contém os scripts utilitários em Python para execução de auditorias e geração de relatórios individuais via linha de comando (CLI), eliminando a necessidade da interface Streamlit para processamentos em lote ou automações.

---

## Pré-requisitos

Certifique-se de que o ambiente virtual (`.venv`) está ativado e com as dependências instaladas:

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

---

## 1. Execução de Auditoria (`run_audit.py`)

O script `run_audit.py` é responsável por rodar todos os procedimentos de auditoria definidos no mapa de verificação contra as bases de dados e fontes de informação carregadas. 

### Principais Saídas:
- **`resultado_auditoria.json`**: Contém o objeto de contexto de todos os auditados serializado em JSON.
- **`tabelas_consolidadas_auditoria.xlsx`**: Pasta de trabalho em Excel com as abas consolidadas:
  - `Achados por Auditado`
  - `Encaminhamentos por Auditado`
  - `Situações Inconformes`
  - `Ranking de Auditados` (classificado por quantidade de achados e situações inconformes)

### Parâmetros:
*   `-a`, `--auditados` (Obrigatório): Caminho para a planilha de cadastro de auditados (ex: `bd_auditados.xlsx`).
*   `-m`, `--mapa` (Obrigatório): Caminho para o mapa de verificação de achados (ex: `mapa-verificacao-achados.xlsx`).
*   `-f`, `--fontes` (Obrigatório, aceita múltiplos arquivos): Caminho para as planilhas Excel que servem como fonte de dados (ex: `20260607-respostas-questionario.xlsx`).
*   `-j`, `--output-json` (Opcional): Caminho personalizado para salvar o JSON (padrão: `.output_reports/resultado_auditoria.json`).
*   `-x`, `--output-xlsx` (Opcional): Caminho personalizado para salvar a planilha de tabelas consolidadas (padrão: `.output_reports/tabelas_consolidadas_auditoria.xlsx`).

### Exemplo de Uso:
```bash
python cli/run_audit.py \
  -a /caminho/para/bd_auditados.xlsx \
  -m /caminho/para/mapa-verificacao-achados.xlsx \
  -f /caminho/para/20260607-respostas-questionario.xlsx
```

---

## 2. Geração de Relatórios Individuais (`generate_reports.py`)

O script `generate_reports.py` consome o JSON de resultados gerado pela auditoria (`resultado_auditoria.json`) e renderiza os relatórios individuais de procedimentos no formato Word (`.docx`) e/ou Markdown (`.md`).

### Parâmetros:
*   `--auditados` (Obrigatório): Caminho para o arquivo JSON contendo o contexto dos auditados (ex: `.output_reports/resultado_auditoria.json`).
*   `--templates` (Obrigatório, aceita múltiplos arquivos): Caminho para os arquivos de template em Markdown/Word (ex: `01-Relatorio_Individual.md`).
*   `--context-files` (Opcional, aceita múltiplos arquivos): Planilhas Excel com informações e variáveis adicionais indexadas pela coluna `sigla`.
*   `--resource-files` (Opcional, aceita múltiplos arquivos): Imagens, diretórios de imagens ou arquivos ZIP contendo as evidências/imagens referenciadas no relatório.
*   `--auditados-select` (Opcional, aceita múltiplos termos): Lista de siglas de auditados específicos a serem gerados (ex: `FTM`). Se omitido, todos os auditados no JSON que foram auditados serão processados.
*   `--output-dir` (Opcional): Diretório de saída para os relatórios (padrão: `.output_reports`).
*   `--reference-docx` (Opcional): Documento Word de referência usado pelo Pandoc para formatação de estilos (padrão: `docs/template-base-estilos-sigiloso.docx`).

### Exemplo de Uso:
```bash
python cli/generate_reports.py \
  --auditados .output_reports/resultado_auditoria.json \
  --templates /caminho/para/01-Relatorio_Individual.md \
  --context-files /caminho/para/iGovTI-2026.xlsx \
  --resource-files /caminho/para/img/ \
  --auditados-select FTM
```

---

## Diretório de Saída e Controle de Versão

Por padrão, ambos os scripts direcionam suas saídas para a pasta oculta **`.output_reports/`**. Esta pasta está configurada no `.gitignore` do projeto para evitar commits acidentais de relatórios temporários ou dados sensíveis gerados localmente.
