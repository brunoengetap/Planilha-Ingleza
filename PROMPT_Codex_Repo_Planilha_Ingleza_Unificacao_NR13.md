# PROMPT PARA CODEX — Melhorar, unificar e padronizar planilhas NR-13 no repositório GitHub

## Contexto do repositório

Você está trabalhando no repositório GitHub:

`brunoengetap/Planilha-Ingleza`

Branch atual:

`main`

Os arquivos `.xlsx` de entrada já estão no diretório raiz do repositório. Não é necessário baixar arquivos externos nem solicitar novos uploads.

Pela estrutura atual do repo, existem planilhas como:

- `Engetap_NR13_ALL-ING-004_2026-06-20 -SETORES VARIADOS VER COLUNA K.xlsx`
- `Engetap_NR13_FCC-ING-003_2026-06-20 - UTILIDADES.xlsx`
- `Engetap_NR13_FCC-ING-TUB-001_2026-06-20 - TUBULAÇÕES.xlsx`
- `Engetap_NR13_FCC_-_ING_-_EMBALAGEM_-_001_2026-06-20 - EMBALAGEM.xlsx`
- `Engetap_NR13_FCC_-_ING_-_PET_-001_2026-06-20 - PET.xlsx`
- `Engetap_NR13_FCC_-_ING_-_TQS_-_001_2026-06-20 - ÁREA DE TANQUES.xlsx`
- `Engetap_NR13_PHF-ING-001_2026-06-20 UTILIDADES.xlsx`
- `Engetap_NR13_PHF_-_ING_-_EMBALAGEM_-_001_2026-06-20 - NÃO ENQUADRADOS - EMBALAGEM.xlsx`
- `Engetap_NR13_PHF_-_ING_-_PET_-_001_2026-06-20 - PET.xlsx`
- `Engetap_NR13_PHF_-_ING_-_TQS_-_001_2026-06-20 - Área de Tanques.xlsx`
- `README.md`

Os nomes podem ter pequenas variações, acentos, espaços, hífens, underscores e letras maiúsculas/minúsculas. Portanto, o script deve localizar os arquivos por padrão, usando glob, e não depender de nomes 100% idênticos.

---

## Objetivo

Criar uma solução dentro do próprio repositório para:

1. Ler todas as planilhas NR-13 do diretório raiz.
2. Consolidar os dados em uma planilha única.
3. Melhorar o aspecto visual da planilha final.
4. Criar abas de uso técnico e auditoria.
5. Preservar os arquivos originais sem qualquer alteração.
6. Criar script rerodável para futuras consolidações.

O resultado final deve ser uma planilha profissional, auditável e útil para análise técnica NR-13.

---

## Arquivos de saída esperados no repositório

Crie a seguinte estrutura:

```text
/scripts
  unificar_nr13.py

/outputs
  Engetap_NR13_Unificado_Profissional.xlsx

/reports
  RELATORIO_UNIFICACAO_NR13.md

README.md
```

Atualize o `README.md` com uma seção curta explicando:

- objetivo do projeto;
- como rodar o script;
- qual arquivo final é gerado;
- quais abas existem na planilha consolidada;
- quais validações são executadas.

Se achar necessário, crie também:

```text
requirements.txt
```

com dependências mínimas, preferencialmente apenas:

```text
openpyxl
```

Evite dependências pesadas sem necessidade.

---

## Restrições obrigatórias

- Não modificar os arquivos `.xlsx` originais.
- Não renomear os arquivos `.xlsx` originais.
- Não apagar nenhum arquivo existente.
- Não sobrescrever manualmente dados técnicos.
- Não inventar informações ausentes.
- Não depender de Microsoft Excel, LibreOffice ou interface gráfica.
- Não usar serviços externos.
- Não usar OCR.
- Não criar solução que dependa de caminho absoluto local.
- O script deve funcionar rodando a partir da raiz do repo.

O script deve ser executável assim:

```bash
python scripts/unificar_nr13.py
```

---

## Detecção dos arquivos de entrada

No script, localizar automaticamente arquivos `.xlsx` no diretório raiz do repo.

Excluir da leitura:

- arquivos dentro de `/outputs`;
- arquivos temporários do Excel começando com `~$`;
- o próprio arquivo final `Engetap_NR13_Unificado_Profissional.xlsx`, caso exista;
- qualquer arquivo que não pareça ser planilha de origem NR-13.

Preferência de filtro:

```python
*.xlsx
```

e depois validar se o arquivo possui aba `NR-13`.

Se a aba `NR-13` não existir, registrar no log e ignorar o arquivo.

---

## Estrutura esperada das planilhas de origem

As planilhas de origem geralmente possuem uma aba chamada:

`NR-13`

Essa aba normalmente contém cerca de 71 colunas, aproximadamente de `A` até `BS`.

O script não deve depender rigidamente do número de colunas, pois pode haver variações entre arquivos.

---

## Abas obrigatórias da planilha final

A planilha final deve conter, no mínimo:

1. `Painel`
2. `Base_Limpa`
3. `Base_Bruta`
4. `Vasos`
5. `Tanques`
6. `Tubulações`
7. `Pendencias`
8. `Dicionario_Campos`
9. `Log_Importacao`

---

# 1. Aba `Base_Bruta`

Aba de rastreabilidade.

Deve conter os dados exatamente como vieram dos arquivos, apenas empilhados.

Adicionar somente as colunas de origem:

- `Fonte_Arquivo`
- `Fonte_Aba`
- `Linha_Origem`

Regras:

- Não corrigir valores.
- Não apagar valores.
- Não substituir placeholders.
- Não normalizar tecnicamente os campos.
- Manter todas as colunas originais encontradas.
- Se houver divergência de cabeçalhos entre arquivos, criar uma união de colunas preservando todos os campos.

---

# 2. Aba `Base_Limpa`

Aba principal de trabalho.

Deve conter os mesmos registros da `Base_Bruta`, mas com colunas adicionais de normalização e auditoria.

Adicionar antes das colunas originais:

- `Fonte_Arquivo`
- `Fonte_Aba`
- `Linha_Origem`
- `Unidade_Arquivo`
- `Area_Arquivo`
- `OS_Normalizada`
- `Tipo_Normalizado`
- `TAG_Normalizada`
- `Data_Coleta_Formatada`
- `Flag_Placeholder_76_77`
- `Colunas_Com_Placeholder`
- `Flag_Dados_Incompletos`
- `Flag_Duplicidade_Potencial`
- `Observacao_Auditoria`

Depois dessas colunas, manter as colunas originais.

---

# 3. Aba `Painel`

Criar uma aba visual de resumo, com layout profissional.

Deve conter indicadores como:

- Total de arquivos processados
- Total de registros importados
- Total por unidade: `FCC`, `PHF`, `ALL`, outras
- Total por tipo: `vaso`, `tanque`, `tubulacao`, `indefinido`
- Total por área/setor
- Total por situação NR-13
- Total por status documental
- Total de registros com pendências de campo
- Total de registros com pendências documentais
- Total de registros sem link do Drive
- Total de registros sem fotos ou com poucas fotos
- Total de registros com placeholders suspeitos `76` ou `77`
- Total de duplicidades potenciais

O painel deve ser legível e bonito, com:

- título grande;
- blocos de KPIs;
- tabelas-resumo;
- largura de colunas ajustada;
- cores discretas e profissionais.

Gráficos simples são desejáveis, mas não obrigatórios se aumentarem muito a complexidade.

---

# 4. Aba `Vasos`

Criar visão filtrada com registros onde:

`Tipo_Normalizado = vaso`

Priorizar colunas como:

- OS
- Cliente
- TAG
- Fabricante
- N_Equip
- Ano_Fab
- Categoria
- Localizacao
- Setor
- PMTA
- P_Trabalho
- P_Teste
- Temperatura
- Fluido
- Classe_Fluido
- Volume
- Diametro
- Comprimento
- Material
- Manometro
- PSV
- Situacao_NR13
- Status_Docs
- Pendencias_Campo
- Pendencias_Documentais
- Qtd_Fotos
- Drive_URL
- Data_Coleta
- Inspetor

Não perder colunas importantes se os nomes estiverem com pequenas variações. Use busca flexível por nome normalizado de cabeçalho.

---

# 5. Aba `Tanques`

Criar visão filtrada com registros onde:

`Tipo_Normalizado = tanque`

Atenção técnica importante:

Tanque metálico não deve ser tratado visualmente como vaso de pressão.

Na aba `Tanques`, não dar destaque principal a:

- PMTA
- pressão de trabalho
- pressão de teste hidrostático

Esses campos podem existir por herança indevida do formulário de vasos. Eles podem permanecer na `Base_Limpa` por rastreabilidade, mas não devem ser campos prioritários na visão de tanques.

Priorizar colunas como:

- OS
- Cliente
- TAG
- Localizacao
- Setor
- Placa_Indelevel
- Necessita_TAG
- Fixado
- Material
- Fluido
- Classe_Fluido
- Volume
- Diametro
- Comprimento
- Aterramento
- Trab_Altura
- Espaco_Confinado
- Enquadra_NR13
- Base_Enquadramento
- Motivo_Nao_Enq
- Situacao_NR13
- Status_Docs
- Disp_Sobrepressao
- Tipo_Disp_Sobrepressao
- Disp_Vacuo
- Valv_Corta_Chamas
- Plano_Manut_Disp
- Cert_Disp_Seguranca
- Pendencias_Campo
- Pendencias_Documentais
- Qtd_Fotos
- Drive_URL
- Data_Coleta
- Inspetor

---

# 6. Aba `Tubulações`

Criar visão filtrada com registros onde:

`Tipo_Normalizado = tubulacao`

Priorizar colunas como:

- OS
- Cliente
- TAG
- Localizacao
- Setor
- Fluido
- Classe_Fluido
- Material
- Diametro
- Comprimento
- Necessita_TH
- Ensaios_END
- Processo
- Enquadra_NR13
- Base_Enquadramento
- Motivo_Nao_Enq
- Situacao_NR13
- Status_Docs
- Pendencias_Campo
- Pendencias_Documentais
- Qtd_Fotos
- Drive_URL
- Data_Coleta
- Inspetor

---

# 7. Aba `Pendencias`

Criar aba com todos os registros que precisam de atenção.

Incluir registros quando qualquer uma das condições abaixo ocorrer:

- `Pendencias_Campo` preenchido;
- `Pendencias_Documentais` preenchido;
- `Status_Docs` indicar documentação ausente, parcial, não localizada ou similar;
- `Flag_Placeholder_76_77 = Sim`;
- `Flag_Dados_Incompletos = Sim`;
- `Flag_Duplicidade_Potencial = Sim`;
- `Qtd_Fotos` vazio, zero ou muito baixo;
- `Drive_URL` vazio.

---

# 8. Aba `Dicionario_Campos`

Criar aba explicando os principais campos.

Colunas:

- `Campo`
- `Descrição`
- `Tipo_Esperado`
- `Exemplo`
- `Aplicável_a`
- `Observações`

Incluir pelo menos os campos técnicos principais e os campos de auditoria adicionados.

---

# 9. Aba `Log_Importacao`

Criar aba técnica com:

- arquivo lido;
- aba lida;
- quantidade de linhas importadas;
- quantidade de colunas encontradas;
- cabeçalhos encontrados;
- se os cabeçalhos bateram com o padrão;
- quantidade de placeholders `76`/`77`;
- quantidade de linhas com placeholders;
- erros ou avisos;
- data/hora da execução.

---

## Regras de normalização

### Cabeçalhos

Criar função para normalizar cabeçalhos internamente:

- remover acentos;
- converter para minúsculas;
- remover quebras de linha;
- remover espaços extras;
- substituir pontuações por `_`;
- colapsar múltiplos `_`.

Mas preservar o nome original da coluna na saída sempre que possível.

### OS

Criar `OS_Normalizada`.

Exemplos:

- `FCC - ING - EMBALAGEM - 001` -> `FCC-ING-EMBALAGEM-001`
- `PHF_-_ING_-_PET_-_001` -> `PHF-ING-PET-001`
- `FCC_-_ING_-_TQS_-_001` -> `FCC-ING-TQS-001`

Não substituir a OS original.

### Unidade

Inferir `Unidade_Arquivo` pelo nome do arquivo e/ou pela OS:

- `FCC`
- `PHF`
- `ALL`
- outros valores encontrados

### Área

Inferir `Area_Arquivo` pelo nome do arquivo, OS ou setor:

- `UTILIDADES`
- `ÁREA DE TANQUES`
- `TUBULAÇÕES`
- `PET`
- `EMBALAGEM`
- `SETORES VARIADOS`
- `NÃO ENQUADRADOS`
- outro valor encontrado

### Tipo

Criar `Tipo_Normalizado` com valores:

- `vaso`
- `tanque`
- `tubulacao`
- `indefinido`

Aceitar variações de acentuação, caixa e grafia.

Exemplos:

- `tubulação`, `tubulacao`, `tubulações` -> `tubulacao`
- `tanque`, `tanques`, `área de tanques`, `tqs` -> `tanque`
- `vaso`, `vasos`, `utilidades`, `pet`, `embalagem` podem precisar de inferência pelos campos e dados existentes

Quando não houver segurança, usar `indefinido` e registrar em `Observacao_Auditoria`.

### TAG

Criar `TAG_Normalizada`:

- remover espaços extras;
- converter para maiúsculas;
- preservar hífens e números;
- deixar vazio se não houver TAG.

### Datas

Criar `Data_Coleta_Formatada` no formato:

`dd/mm/aaaa hh:mm`

Preservar o valor original.

### Sim/Não

Padronizar variações para:

- `Sim`
- `Não`

Sem substituir na `Base_Bruta`.

---

## Tratamento dos placeholders 76 e 77

Valores `76` e `77` aparecem como suspeitos em algumas planilhas.

Regra obrigatória:

- Na `Base_Bruta`, preservar exatamente como vieram.
- Na `Base_Limpa`, detectar esses valores e marcar auditoria.
- Não apagar automaticamente sem deixar rastro.
- Criar `Flag_Placeholder_76_77 = Sim` quando a linha tiver `76` ou `77`.
- Criar `Colunas_Com_Placeholder` com a lista das colunas afetadas.
- Em colunas textuais ou campos onde `76`/`77` claramente não fazem sentido, substituir por vazio ou `Pendente de conferência`, mas somente na `Base_Limpa`.
- Registrar no `Log_Importacao` a quantidade por arquivo.
- Explicar no relatório `.md`.

Importante:

Se `76` ou `77` puderem ser valor técnico plausível em uma coluna numérica, não apagar. Apenas marcar.

---

## Duplicidades potenciais

Marcar `Flag_Duplicidade_Potencial = Sim` quando houver repetição de:

- `OS_Normalizada + TAG_Normalizada`

ou

- `Cliente + TAG_Normalizada + Tipo_Normalizado`

Não excluir duplicidades automaticamente.

Registrar no relatório.

---

## Links do Drive

Se houver coluna equivalente a `Drive_URL`, transformar os valores em hiperlinks clicáveis na planilha final.

Não alterar o endereço original.

---

## Regras visuais

Aplicar formatação profissional em todas as abas:

- congelar primeira linha;
- aplicar filtro automático;
- usar tabelas do Excel quando possível;
- cabeçalho azul escuro, fonte branca, negrito;
- bordas leves;
- alinhamento vertical centralizado;
- quebra de texto em campos longos;
- largura de colunas ajustada;
- limite de largura máxima para campos longos;
- datas formatadas;
- hiperlinks clicáveis;
- destaque visual para pendências;
- destaque visual para placeholders;
- destaque visual para duplicidade potencial.

No `Painel`, aplicar visual de dashboard:

- título no topo;
- blocos de indicadores;
- tabelas de resumo;
- cores discretas;
- layout limpo.

---

## Relatório obrigatório

Criar:

`reports/RELATORIO_UNIFICACAO_NR13.md`

O relatório deve conter:

1. Objetivo da consolidação.
2. Arquivos processados.
3. Arquivos ignorados e motivo.
4. Quantidade de registros por arquivo.
5. Quantidade total de registros.
6. Quantidade por tipo.
7. Quantidade por unidade.
8. Quantidade de registros com placeholders `76`/`77`.
9. Quantidade de registros com pendências.
10. Quantidade de duplicidades potenciais.
11. Regras de normalização aplicadas.
12. Limitações conhecidas.
13. Como rerodar o script.

---

## Validações obrigatórias

Antes de finalizar, rodar validações no script:

1. `Base_Bruta` e `Base_Limpa` devem ter a mesma quantidade de registros.
2. Todos os arquivos `.xlsx` válidos com aba `NR-13` devem ser processados.
3. Todas as linhas lidas devem ter `Fonte_Arquivo` preenchido.
4. Todas as linhas lidas devem ter `Linha_Origem` preenchida.
5. O arquivo final deve ser carregável por `openpyxl.load_workbook`.
6. Nenhum arquivo original pode ser alterado.
7. O script deve conseguir rodar novamente sem quebrar.
8. Os diretórios `/outputs` e `/reports` devem ser criados automaticamente se não existirem.
9. O relatório `.md` deve ser gerado mesmo que algum arquivo tenha alerta.
10. O log de importação deve registrar avisos em vez de interromper toda a execução quando houver problema isolado.

---

## Implementação sugerida

Use Python com `openpyxl`.

Estrutura recomendada do script:

```python
from pathlib import Path
import re
import unicodedata
from datetime import datetime
from openpyxl import load_workbook, Workbook
```

Funções sugeridas:

- `repo_root()`
- `find_input_workbooks(root)`
- `read_nr13_sheet(path)`
- `normalize_header(value)`
- `normalize_os(value)`
- `normalize_text(value)`
- `infer_unidade(filename, row)`
- `infer_area(filename, row)`
- `normalize_tipo(filename, row)`
- `normalize_tag(value)`
- `detect_placeholders(row)`
- `detect_incomplete_data(row)`
- `detect_duplicates(rows)`
- `build_base_bruta()`
- `build_base_limpa()`
- `build_type_views()`
- `build_pendencias()`
- `build_dashboard()`
- `build_dictionary()`
- `build_import_log()`
- `apply_workbook_styles()`
- `write_report()`
- `validate_output()`

---

## Critério de aceite

A implementação será considerada concluída quando:

- `python scripts/unificar_nr13.py` rodar sem erro;
- `outputs/Engetap_NR13_Unificado_Profissional.xlsx` for gerado;
- `reports/RELATORIO_UNIFICACAO_NR13.md` for gerado;
- `README.md` for atualizado;
- os arquivos originais permanecerem intactos;
- a planilha final abrir normalmente no Excel;
- a planilha final tiver as abas obrigatórias;
- a `Base_Bruta` preservar os dados originais;
- a `Base_Limpa` tiver normalizações e flags de auditoria;
- o `Painel` estiver visualmente melhor que as planilhas originais;
- o `Log_Importacao` mostrar o que foi processado;
- o relatório explicar claramente os alertas.

---

## Resposta final esperada do Codex

Ao terminar, responda com:

1. Arquivos criados/alterados.
2. Comando para rodar.
3. Quantidade de planilhas processadas.
4. Quantidade total de registros consolidados.
5. Principais alertas encontrados.
6. Confirmação de que os arquivos originais não foram modificados.
