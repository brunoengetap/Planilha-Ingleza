# Planilha-Ingleza

Projeto para unificar planilhas técnicas NR-13 da Engetap em uma pasta de trabalho única, auditável e pronta para análise.

## Como rodar

A partir da raiz do repositório, execute:

```bash
python scripts/unificar_nr13.py
```

O script localiza automaticamente as planilhas `.xlsx` NR-13 na raiz do repositório, ignora arquivos temporários, ignora saídas em `outputs/` e processa somente arquivos com aba `NR-13`. Instale as dependências e gere a planilha localmente com:

```bash
python -m pip install -r requirements.txt
python scripts/unificar_nr13.py
```

## Arquivo gerado

A consolidação é gravada em:

```text
outputs/Engetap_NR13_Unificado_Profissional.xlsx
```

Esse arquivo `.xlsx` final **não fica versionado no repositório**. Ele deve ser gerado localmente pelo script para evitar depender do download de binários pela interface do Codex/GitHub.

Também é gerado o relatório:

```text
reports/RELATORIO_UNIFICACAO_NR13.md
```

## Abas da planilha consolidada

- `Painel`: resumo executivo com KPIs por arquivo, tipo, unidade, área, pendências, placeholders e duplicidades.
- `Base_Bruta`: dados empilhados exatamente como lidos, com `Fonte_Arquivo`, `Fonte_Aba` e `Linha_Origem`.
- `Base_Limpa`: base de trabalho com colunas auxiliares de normalização e auditoria.
- `Vasos`: visão filtrada de vasos de pressão.
- `Tanques`: visão filtrada de tanques, sem priorizar campos típicos de vasos como PMTA.
- `Tubulações`: visão filtrada de tubulações.
- `Pendencias`: registros que exigem atenção técnica ou documental.
- `Dicionario_Campos`: descrição dos principais campos técnicos e de auditoria.
- `Log_Importacao`: rastreabilidade da execução por arquivo.

## Validações executadas

O script valida que `Base_Bruta` e `Base_Limpa` tenham a mesma quantidade de registros, que toda linha importada tenha arquivo e linha de origem, que diretórios de saída sejam criados automaticamente e que os arquivos originais da raiz permaneçam sem alteração. Valores suspeitos `76` e `77`, dados incompletos e duplicidades potenciais são sinalizados sem apagar dados brutos.

## Dependências

O repositório inclui `requirements.txt` com `openpyxl` para validação externa da planilha gerada. A geração principal usa apenas bibliotecas padrão do Python para evitar dependências pesadas.
