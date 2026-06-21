# Relatório de Unificação NR-13

Gerado em: 21/06/2026 16:18:51.

## Objetivo
Consolidar as planilhas NR-13 da raiz do repositório em uma base única, auditável e rerodável.

## Arquivos processados
- Engetap_NR13_ALL-ING-004_2026-06-20 -SETORES VARIADOS VER COLUNA K.xlsx: 29 registros, 71 colunas.
- Engetap_NR13_FCC-ING-003_2026-06-20 - UTILIDADES.xlsx: 10 registros, 71 colunas.
- Engetap_NR13_FCC-ING-TUB-001_2026-06-20 - TUBULAÇÕES.xlsx: 8 registros, 71 colunas.
- Engetap_NR13_FCC_-_ING_-_EMBALAGEM_-_001_2026-06-20 - EMBALAGEM.xlsx: 17 registros, 71 colunas.
- Engetap_NR13_FCC_-_ING_-_PET_-001_2026-06-20 - PET.xlsx: 3 registros, 71 colunas.
- Engetap_NR13_FCC_-_ING_-_TQS_-_001_2026-06-20 - ÁREA DE TANQUES.xlsx: 5 registros, 71 colunas.
- Engetap_NR13_PHF-ING-001_2026-06-20 UTILIDADES.xlsx: 7 registros, 71 colunas.
- Engetap_NR13_PHF_-_ING_-_EMBALAGEM_-_001_2026-06-20 - NÃO ENQUADRADOS - EMBALAGEM.xlsx: 22 registros, 71 colunas.
- Engetap_NR13_PHF_-_ING_-_PET_-_001_2026-06-20 - PET.xlsx: 18 registros, 71 colunas.
- Engetap_NR13_PHF_-_ING_-_TQS_-_001_2026-06-20 - Área de Tanques.xlsx: 6 registros, 71 colunas.

## Arquivos ignorados
- Nenhum.

## Totais
- Planilhas processadas: 10
- Registros consolidados: 125
- Registros com placeholders 76/77: 0
- Registros com pendências: 84
- Duplicidades potenciais: 4

## Quantidade por tipo
- vaso: 90
- indefinido: 16
- tanque: 11
- tubulacao: 8

## Quantidade por unidade
- PHF: 53
- FCC: 43
- ALL: 29

## Regras de normalização aplicadas
- Cabeçalhos normalizados internamente para busca flexível, preservando nomes originais nas abas de dados.
- OS, unidade, área, tipo, TAG e datas recebem colunas auxiliares sem alterar a base bruta.
- Valores 76/77 são preservados na Base_Bruta e marcados na Base_Limpa.
- Duplicidades potenciais são marcadas, nunca excluídas.

## Limitações conhecidas
- Inferências dependem dos nomes de arquivos, cabeçalhos e conteúdo disponível; campos ausentes não são inventados.
- O ambiente atual não tinha openpyxl instalado via rede, portanto o gerador usa OOXML padrão com biblioteca nativa do Python.

## Como rerodar
```bash
python scripts/unificar_nr13.py
```

Arquivos originais preservados: Sim.
