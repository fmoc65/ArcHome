# Contexto do projeto R3Integrador

Última atualização: 12/08/2026.

Este arquivo é o ponto de retomada do trabalho realizado no projeto. Ele reúne
as regras de negócio confirmadas, fontes de dados, decisões de implementação,
estado do SQLite, arquivos gerados e pendências conhecidas.

## 1. Projeto e diretórios

- Projeto: `/home/fernando/Projetos/Work/ARCHOME/ArcHome/R3Integrador`
- Solução: `R3Integrador.slnx`
- Banco local: `R3IntegradorDb.db`
- Saídas: `Saida/`
- Planilhas externas: `/home/fernando/Projetos/Work/ARCHOME/Planilhas`
- Aplicação principal: `src/R3Integrador.Web`
- Console/fluxo legado: `src/R3Integrador.Console`
- Layout completo de importação do ERP: 60 colunas, aba `IMPORTACAO_ERP`
- Layout resumido de atualização: marca, referência, preço de venda, preço de
  fábrica, descrição, unidade fabril, modelo e cor.

Comandos de compilação e execução:

```bash
dotnet build R3Integrador.slnx --no-restore -m:1
dotnet run --project src/R3Integrador.Web/R3Integrador.Web.csproj
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj
```

O parâmetro `-m:1` evita uma falha silenciosa de concorrência observada no
MSBuild deste ambiente. A solução compilou em 10/08/2026 com zero erros e zero
avisos.

## 2. Banco SQLite

O banco `R3IntegradorDb.db` possui:

- `Produtos`: as 60 colunas do layout ERP e a coluna adicional
  `TabelaOrigem`, usada para identificar a procedência de cada registro.
- `Usuario`: `Id`, `Nome`, `Login`, `Senha` e `Ativo`. `Ativo` é inteiro
  booleano (`0`/`1`) e seu padrão é `1` (`true`).
- `VillaColProdutosFonte`: dados normalizados extraídos da tabela/PDF VillaCol.

Contagens verificadas em 10/08/2026:

| TabelaOrigem | Registros |
|---|---:|
| `VAREJO` | 247 |
| `ATLAS_REVENDA_35` | 203 |
| `VILLACOL_REVENDA` | 63 |
| `VILLACOL_DEL5` | 10 |
| `VILLACOL_DEL10` | 10 |
| `VILLACOL_DEL15` | 10 |
| `VILLACOL_DEL20` | 10 |
| `VILLACOL_DEL25` | 10 |
| `VILLACOL_DEL30` | 10 |

Ao reimportar uma origem, somente os produtos daquela `TabelaOrigem` devem ser
substituídos. Os pipelines Atlas e VillaCol geram backup do SQLite antes da
carga.

Criação/importação genérica do banco:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --criar-banco /caminho/IMPORTACAO_ERP.xlsx R3IntegradorDb.db VAREJO
```

Implementação principal:

- `src/R3Integrador.Infrastructure/Persistence/R3IntegradorDbInitializer.cs`
- `src/R3Integrador.Infrastructure/Persistence/AtualizacaoPrecoFabricaService.cs`

## 3. Regras gerais confirmadas

### 3.1 Produtos vendidos em metro quadrado

- Se a unidade de venda for `M2`, a quantidade de embalagem de venda deve ser
  `0`.
- O comprador compra por metro quadrado; `m²/caixa`, peças por caixa e peças por
  pallet são informações logísticas, não quantidade mínima/embalagem de venda.
- Para itens que não são `M2`, o fluxo atual normalmente usa embalagem de venda
  `1`, salvo regra expressa da tabela.
- Foi discutido um caso de caixa com 5 unidades e pallet com 40 unidades. Esses
  valores não devem substituir a unidade comercial em `M2`.

### 3.2 Atualização de preço de fábrica

- Ao atualizar produtos já existentes, utilizar somente o preço de fábrica da
  nova tabela.
- Não usar preço Del Credere como preço de fábrica.
- Preservar os demais dados cadastrais e fiscais dos produtos existentes.
- Separar o resultado em:
  1. planilha de atualização dos itens já cadastrados;
  2. planilha de inclusão/complementação fiscal dos itens ausentes.
- A planilha de inclusão deve destacar os campos fiscais que ainda precisam ser
  informados.

Comando existente:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --atualizar-preco-fabrica /caminho/TABELA_VAREJO.xlsx R3IntegradorDb.db Saida
```

Saídas existentes desse fluxo:

- `Saida/IMPORTACAO_ERP_ATUALIZACAO_PRECO_FABRICA_VAREJO_ATUALIZADO_20260805_205956.xlsx`
- `Saida/IMPORTACAO_ERP_INCLUSAO_VAREJO_PENDENTE_FISCAL_20260805_205956.xlsx`

### 3.3 Del Credere

- Nas tabelas Del Credere, somar ao preço-base:
  - IPI: `0,65%`;
  - taxa de cartão: `4,71%`;
  - acréscimo total aditivo: `5,36%`.
- Objetivo: o orçamento já sair com o preço total.
- A atualização resumida não deve recriar importações de produtos já
  importados. Ela serve somente para atualizar preços.
- No fluxo solicitado, gerar somente Villagres.
- Gerar seis tabelas Villagres: `DEL5`, `DEL10`, `DEL15`, `DEL20`, `DEL25` e
  `DEL30`.
- Não incluir Villa Art nesse fluxo e não gerar novamente o layout completo de
  importação ERP.

Menu existente:

- opção `16 - Atualização Del Credere Villagres (layout resumido)`.

Saídas atuais:

- pasta `Saida/ATUALIZACAO_DELCREDERE_VILLAGRES_20260808_104527/`;
- contém as seis planilhas `ATUALIZACAO_DELCREDERE_DEL*_VILLAGRES.xlsx`.

Layout resumido de referência fornecido pelo usuário:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/layout_resumida.xlsx`
  (confirmar a presença do arquivo antes de uma nova execução).

### 3.4 Fiscal Villagres/Varejo

Premissas informadas para os produtos Villagres/Varejo que estavam em processo
de inclusão:

- NCM: `69041000` para todos os produtos desse conjunto;
- origem: `SP`;
- empresa ARCHOME: Simples Nacional;
- CSOSN confirmado posteriormente: `500`.

Essas premissas não devem ser aplicadas automaticamente a outros fornecedores,
como VillaCol ou Nina, porque cada fluxo possui regras próprias.

## 4. VillaCol

Diretório de origem:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/VillaCol`

Fontes:

- seis planilhas Del Credere aprovadas pelo contador, de `DEL5` a `DEL30`;
- PDF `TABELA VAREJO 04-05-2026.pdf` com os produtos de revenda.

Decisão de escopo importante:

- A tabela de revenda deve conter todos os itens extraídos do PDF.
- Não assumir que todos os itens de revenda pertencem a todas as tabelas Del
  Credere.
- Cada planilha `DEL5` a `DEL30` permanece somente com os itens que já existiam
  nela antes da correção.
- Resultado final: 63 SKUs na revenda e 10 itens históricos em cada Del
  Credere.

Tratamentos realizados:

- referências técnicas criadas/vinculadas onde os códigos estavam ausentes;
- análise da `DEL30`, considerada consistente após auditoria;
- `Grupo`, `Subgrupo` e `Observação` higienizados, sem `NaN`;
- preço de venda forçado para `0` no pipeline VillaCol;
- marca preenchida como `VillaCol`;
- informações fiscais herdadas das planilhas aprovadas pelo contador;
- fontes originais não sobrescritas.

Fiscal normalizado no relatório da revisão 2:

| Campo | Valor |
|---|---|
| NCM | `32149000` |
| UF origem | `SP` |
| CST | `000` |
| COFINS CST | `49` |
| IPI CST | `99` |
| PIS CST | `49` |
| CSOSN | `102` |
| CFOP dentro | `5102` |
| CFOP fora | `6102` |
| Enquadramento IPI | `999` |
| PIS origem | `0` |
| COFINS origem | `0` |
| IBS | `0.1` |
| CBS | `0.9` |
| Classificação tributária | `000001` |

Pipeline:

- `scripts/villacol_pipeline.py`
- dependências em `requirements-villacol.txt`
- ambiente Python usado: `.venv-villacol`

Execução:

```bash
.venv-villacol/bin/python scripts/villacol_pipeline.py \
  --input-dir /home/fernando/Projetos/Work/ARCHOME/Planilhas/VillaCol \
  --database R3IntegradorDb.db
```

Resultado definitivo mais recente:

- `Saida/VILLACOL_PROCESSADO_20260810_REVISAO_2/`
- seis importações Del Credere corrigidas;
- `IMPORTACAO_ERP_REVENDA_VILLACOL.xlsx` com 63 itens;
- `ITENS_EXTRAIDOS_PDF_VILLACOL.csv`;
- `RELATORIO_AUDITORIA_VILLACOL.json`;
- backup `R3IntegradorDb_ANTES_VILLACOL.db.bak`.

QA da revisão 2:

- 63 itens extraídos do PDF;
- 63 itens na importação de revenda;
- seis arquivos Del Credere corrigidos;
- 10 itens em cada Del Credere;
- códigos de fábrica únicos;
- preço de venda igual a zero;
- totais do PDF, planilhas e SQLite conferidos.

## 5. Atlas

Diretório de origem:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/Atlas`

Fontes:

- `IMPORTACAO_ERP_ATLAS_REVENDA_35_PROVISORIA_CONTADOR_20260717_OK_CONTADOR.xlsx`
- `TABELA ATLAS REVENDA 35% - MAIO 2026.xlsx`

Pipeline:

- `scripts/atlas_audit_import.py`

Tratamentos e resultado:

- 203 produtos na origem e 203 na importação;
- nenhuma divergência de produto ou preço;
- campos fiscais preservados da planilha aprovada pelo contador;
- observações vazias;
- 194 produtos em `M2` corrigidos para embalagem de venda `0`;
- 9 produtos em `PC` mantidos com embalagem de venda `1`;
- preços de revenda não foram recalculados porque já conferiam;
- origem `ATLAS_REVENDA_35` substituída no SQLite, com 203 registros.

Saída definitiva:

- `Saida/ATLAS_PROCESSADO_20260810/IMPORTACAO_ERP_ATLAS_REVENDA_35_CORRIGIDA.xlsx`
- relatório `Saida/ATLAS_PROCESSADO_20260810/RELATORIO_AUDITORIA_ATLAS.json`
- backup `Saida/ATLAS_PROCESSADO_20260810/R3IntegradorDb_ANTES_ATLAS.db.bak`

## 6. Derosso

Foram separados os fluxos de revenda e representação. Saídas existentes:

- `Saida/IMPORTACAO_ERP_DEROSSO_REVENDA_PROVISORIA_20260808_123810.xlsx`
- `Saida/IMPORTACAO_ERP_DEROSSO_REPRESENTACAO_PROVISORIA_20260808_123816.xlsx`

Esses arquivos continuam identificados como provisórios e devem passar pelas
validações fiscais/comerciais aplicáveis antes da importação definitiva.

## 7. Nina Martinelli

### 7.1 Origem e menu

Arquivo localizado e processado:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/Nina/CONVERTIDA_NINA_MARTINELLI_2026_REPRESENTACAO_REV02.xlsx`
- aba de origem: `COLECAO_COMPLETA`;
- 1.141 linhas de produtos;
- menu do console: opção `13 - Processar Planilha NINA MARTINELLI
  (PROVISORIA)`.

O leitor também aceita uma planilha que já contenha a aba `IMPORTACAO_ERP`.

### 7.2 Premissas fiscais Nina

Todas as alíquotas são armazenadas em C# como `decimal` e como frações, não
como percentuais inteiros:

| Parâmetro | Valor decimal |
|---|---:|
| Regime da empresa | Simples Nacional |
| UF da empresa | `SP` |
| UF da fábrica | `SP` |
| ICMS origem | `0.12m` |
| ICMS saída | `0.12m` |
| MVA | `0.81m` |
| IPI | `0.0065m` |
| PIS origem | `0.0165m` |
| COFINS origem | `0.076m` |
| ST esperada quando aplicável | `0.0986m` |

Implementação:

- `src/R3Integrador.Application/Services/NinaFiscalParameters.cs`
- `src/R3Integrador.Application/Services/ICsosnService.cs`
- `src/R3Integrador.Application/Services/CsosnService.cs`
- `src/R3Integrador.Infrastructure/Repositories/NinaMartinelliReaderService.cs`
- registro de dependências no console e no projeto web.

### 7.3 Regra CSOSN/ST Nina

- Operação não tributada indicada explicitamente: CSOSN `400`, ST `0`.
- ST já recolhida/retida/cobrada anteriormente: CSOSN `500`, ST `0.0986`.
- Sem incidência de ST e sem permissão de crédito: CSOSN `102`, ST `0`.
- Sem incidência de ST com permissão de crédito: CSOSN `101`, ST `0`.
- ST de 9,86% cobrada na operação, sem permissão de crédito: CSOSN `202`.
- ST de 9,86% cobrada na operação, com permissão de crédito: CSOSN `201`.
- O serviço aceita `9.86` ou `0.0986`, normaliza e valida contra `0.0986m`.
- Um percentual diferente de 9,86% gera erro de validação, em vez de ser
  aceito silenciosamente.
- Sem indicação explícita de ST, não presumir ST: usar `0` e CSOSN `102`.

A tabela Nina atual não possui coluna nem indicação textual de ST. Portanto,
todos os 1.141 itens da saída atual ficaram com CSOSN `102` e ST `0`.

Referência oficial usada para a semântica dos códigos:

- Portal Nacional da NF-e, tabela CSOSN:
  `https://www.nfe.fazenda.gov.br/Portal/perguntasFrequentes.aspx?AspxAutoDetectCookieSupport=1&tipoConteudo=S%2FEAGUrzRyk%3D`

### 7.4 Mapeamento e saída Nina

- Marca fixa: `NINA MARTINELLI`.
- Grupo: `REVESTIMENTOS`.
- Subgrupo: aplicação da tabela; se ausente, `REVESTIMENTOS`.
- Unidade normalizada: `M2`, `PC` ou `LT`.
- Para `M2`, embalagem de venda `0`.
- O preço de representação da origem foi preservado, pelo comportamento atual,
  tanto como preço de venda quanto como preço de fábrica.
- Quatro referências técnicas determinísticas foram criadas para códigos
  repetidos ou para a descrição `depende do raio`.
- Todas as 1.141 referências da saída são únicas.

Saída definitiva mais recente:

- `Saida/IMPORTACAO_ERP_NINA_MARTINELLI_IMPOSTOS_CONTADOR_20260810_215759.xlsx`

QA da saída:

- 1.141 registros;
- 60 colunas;
- XLSX íntegro;
- nenhuma fórmula;
- nenhuma referência duplicada;
- 736 produtos `M2`, todos com embalagem de venda `0`;
- 396 produtos `PC` e 9 produtos `LT`, com embalagem de venda `1`;
- todas as linhas com UF `SP`, IPI `0.0065`, ICMS origem/saída `0.12`, MVA
  `0.81`, PIS origem `0.0165`, COFINS origem `0.076`, ST `0` e CSOSN `102`.

Pendências Nina:

- NCM, CST, CFOP dentro e CFOP fora permanecem vazios porque não foram
  fornecidos para Nina.
- Não inventar esses campos; solicitar/homologar antes de considerar a planilha
  fiscalmente completa.
- O menu ainda chama o fluxo de `PROVISORIA`, coerente com essas pendências.

## 8. Arquivos e saídas que não devem ser confundidos

- `IMPORTACAO_ERP_*`: layout completo para inclusão/importação ERP.
- `ATUALIZACAO_DELCREDERE_*`: layout resumido, somente atualização de preços.
- Não gerar novamente importações de itens que já foram importados quando o
  pedido for somente atualização.
- A saída mais recente deve ser preferida quando houver duas execuções do mesmo
  fornecedor com timestamps diferentes.
- Para VillaCol, usar `VILLACOL_PROCESSADO_20260810_REVISAO_2`, não a primeira
  execução sem o sufixo de revisão.
- Para Nina, usar o arquivo com timestamp `20260810_215759`, não o teste anterior
  `20260810_215507`.

## 9. Princípios operacionais para próximas alterações

- Não sobrescrever planilhas originais aprovadas pelo contador.
- Auditar contagem, códigos, preços, unidade, embalagem, observação e campos
  fiscais antes de carregar no SQLite.
- Usar `decimal` em cálculos monetários e fiscais em C#.
- Não propagar regras fiscais de um fornecedor para outro sem confirmação.
- Criar backup do banco antes de substituir uma origem material.
- Manter `TabelaOrigem` consistente para permitir substituição isolada.
- Gerar relatório de QA quando houver extração de PDF ou reconciliação entre
  duas fontes.
- Em qualquer divergência entre preço por peça, caixa, pallet ou metro quadrado,
  preservar o valor original e interromper a transformação comercial até a
  unidade correta ser confirmada.

## 9.1 Adama — importação realizada e correção de preços em 17/08/2026

A importação inicial da Adama já foi realizada no ERP usando a planilha
validada pelo contador. Portanto, não gerar uma nova `IMPORTACAO_ERP_ADAMA` para
corrigir preços; produtos divergentes devem ir exclusivamente para atualização
resumida.

Regra comercial confirmada pelo cliente:

- Adama é representação;
- preço de venda e preço de fábrica devem ser iguais à coluna `PREÇO 2026` da
  tabela do fornecedor;
- não aplicar margem, acréscimo ou conversão para `VALOR M²`;
- exemplo confirmado: referências `AAP0001...`, de R$ 254,40 para R$ 63,60.

Auditoria realizada:

- 889 referências únicas na tabela e na importação aprovada;
- 163 produtos haviam recebido `VALOR M²` no lugar de `PREÇO 2026`;
- os 163 preços foram corrigidos;
- a origem `ADAMA` foi sincronizada no `R3IntegradorDb.db` com 889 produtos;
- preço de venda e preço de fábrica não possuem divergências no SQLite.

Para o ERP, usar somente o layout resumido de oito colunas já adotado nas
atualizações: `MARCA`, `REFERENCIA`, `PREÇO VENDA`, `PREÇO DE FÁBRICA`,
`DESCRICAO`, `UNID FABRIL`, `MODELO` e `COR`. A planilha completa corrigida é
apenas base interna de auditoria e banco, não uma nova importação.

## 9.2 Varejo Villagres — recálculo confirmado em 17/08/2026

Para a tabela `001 TABELA VAREJO SP 01.05.26 - ARC HOME.xlsx`, usar a coluna
`DESCONTO` como base e calcular:

`preço final = DESCONTO × (1 + IPI / 100 + 9,86 / 100)`

- IPI de 0,65% para NCM `69072100`;
- IPI zero para a referência `120003`, cujo NCM é `69072200`;
- arredondamento monetário para duas casas;
- não somar o antigo valor fixo de R$ 1,50;
- não aplicar o fator 1,75: o markup será incluído posteriormente pelo cliente;
- `PREÇO VENDA` e `PREÇO DE FÁBRICA` devem receber o mesmo preço final.

A atualização deve usar o layout reduzido de oito colunas. A inclusão mantém o
layout completo. A execução validada está em
`Saida/VAREJO_RECALCULADO_20260817_200225`. Ela preserva, para conferência, a
separação histórica de 245 atualizações e cinco inclusões, cobrindo as 250
referências da fonte. Entretanto, as cinco referências da antiga inclusão
(`120001`, `120003`, `120004`, `120005`, `120006`) já foram confirmadas no ERP
e sincronizadas no SQLite. Portanto, o arquivo operacional recomendado é
`ATUALIZACAO_VAREJO_PRECOS_RECALCULADOS_250_ITENS.xlsx`; não reenviar os cinco
itens como inclusão no mesmo ERP.

## 9.3 Villa Art — recálculo confirmado em 17/08/2026

Para `002 TABELA VILLA ART SP 01.05.26 - ARC HOME.xlsx`, aplicar a mesma regra
comercial confirmada para a tabela 001:

`preço final = DESCONTO × (1 + 0,65 / 100 + 9,86 / 100)`

- usar o preço `DESCONTO` exibido com duas casas como base;
- não usar a coluna `PREÇO SUGERIDO / NA PONTA` como preço de venda;
- não aplicar markup;
- `PREÇO VENDA` e `PREÇO DE FÁBRICA` devem ser iguais ao preço final;
- arredondar para duas casas.

O processamento anterior copiava `DESCONTO` para fábrica e `NA PONTA` para
venda. Na fonte, `NA PONTA` equivale a aproximadamente 1,705063 vezes o preço
com desconto, por isso os valores de venda estavam elevados.

Execução validada: `Saida/VILLA_ART_RECALCULADO_20260817_203923`. A fonte tem
25 referências: 24 já cadastradas geraram atualização reduzida e a referência
nova `123052` gerou uma inclusão completa. Os 25 preços foram validados pela
nova fórmula e a inclusão foi conferida com IPI 0,65%, ST 9,86% e CST de
PIS/COFINS `01`.

## 9.4 Vinílico e rodapé — tabela 003 em 17/08/2026

Fonte: `003 TABELA VINÍLICO SP 01.05.26 - ARC HOME.xlsx`, cuja aba se chama
`VAREJO` apesar de conter vinílicos e rodapés. O leitor dedicado aceita essa
aba alternativa e remove o marcador `***` das duas referências SPC.

Regra de preço dos vinílicos, coerente com o fiscal já existente no mapper:

`preço final = DESCONTO × (1 + 7,92 / 100)`

- IPI zero;
- ST 7,92%;
- sem o antigo valor fixo de R$ 1,50;
- sem markup de 1,75;
- preço de venda igual ao preço de fábrica;
- cálculo sobre o valor `DESCONTO` exibido com duas casas.

A fonte possui 29 referências: 16 vinílicos (`SPC`/`LVT`) e 13 rodapés (`RP`).
Nenhuma consta no SQLite, portanto não foram gerados arquivos vazios de
atualização. A saída validada está em
`Saida/VINILICO_RODAPE_PROCESSADO_20260817_205546` e contém duas importações
separadas.

A importação de vinílicos está fiscalmente preenchida com NCM `39181000`, ST
7,92% e unidade `M2`. A importação de rodapés foi separada com grupo `RODAPE`,
mas permanece identificada como pendente de unidade, NCM e tributação: a fonte
não fornece esses dados e não se deve reutilizar automaticamente o NCM de piso
vinílico para rodapé de poliestireno. Os preços do rodapé foram calculados pela
mesma taxa comercial provisória de 7,92%, mas o arquivo não deve ser carregado
até homologação fiscal e da unidade.

### Validação do contador e carga dos vinílicos — 17/08/2026

A planilha
`/home/fernando/Projetos/Work/ARCHOME/Planilhas/IMPORTACAO_ERP_VINILICO_20260701_200706_CONTADOR.xlsx`
possui os 16 vinílicos e um único perfil fiscal completo. Os campos obrigatórios
foram validados sem faltas, com NCM `39181000`, CST `010`, CST de PIS/COFINS
`01`, ST 7,92%, unidade `M2` e unidade fabril `CX`.

Quatro referências mudaram entre a versão revisada pelo contador e a tabela
003 atual: `LVT96350001` a `LVT96350004` passaram a `LVT94350001` a
`LVT94350004`. Como o perfil fiscal é idêntico nas 16 linhas do contador, ele
foi aplicado aos códigos atuais, mantendo descrições e preços da tabela 003.

Saída definitiva:
`Saida/VINILICO_RODAPE_PROCESSADO_20260817_205546/IMPORTACAO_ERP_INCLUSAO_VINILICO_CONTADOR_VALIDADO.xlsx`.
Ela contém 16 referências, venda igual à fábrica e nenhum campo obrigatório
ausente. A origem `VINILICO` foi importada no SQLite: o banco passou de 1.511
para 1.527 produtos. Nenhum rodapé foi carregado. Backup anterior:
`R3IntegradorDb_ANTES_VINILICO_20260817_210504.db.bak`.

## 10. Atualização Del Credere Villagres — 11/08/2026

Fonte processada:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/delcredere/100 TABELA DEL CREDERE SP ATUALIZADA.xlsx`
- aba: `100 TABELA COM DEL CREDERE`.

O `DelcredereReaderService` foi ajustado para aceitar a aba histórica
`COM DEL CREDERE` e essa nova variação de nome, validando o cabeçalho das seis
faixas de preço antes da leitura.

### Segmentação por cor em B/C

As linhas 318 a 322 da fonte são uma legenda visual; não são produtos. A
segmentação é identificada pelas cores aplicadas nas colunas B e C de cada
referência:

| Segmento | Referências |
|---|---:|
| Villa Premium | 28 |
| Villa Exclusive | 39 |
| Villa Max | 14 |
| Villa Style | 12 |
| Exclusivo Villa Art | 24 |
| Sem segmentação visual | 154 |

O layout resumido não possui uma coluna própria para essa classificação. Para
não perder a informação e sem alterar a marca ou a cor real do produto, os 93
itens Villagres segmentados receberam o sufixo `SEGMENTO: ...` na coluna
`DESCRICAO`. As 24 referências `EXCLUSIVO VILLA ART` foram identificadas, mas
foram excluídas porque este fluxo é exclusivamente Villagres.

### Saída e auditoria

Pasta definitiva desta execução:

- `Saida/ATUALIZACAO_DELCREDERE_VILLAGRES_20260811_191938/`

Arquivos gerados:

- `ATUALIZACAO_DELCREDERE_DEL5_VILLAGRES.xlsx`
- `ATUALIZACAO_DELCREDERE_DEL10_VILLAGRES.xlsx`
- `ATUALIZACAO_DELCREDERE_DEL15_VILLAGRES.xlsx`
- `ATUALIZACAO_DELCREDERE_DEL20_VILLAGRES.xlsx`
- `ATUALIZACAO_DELCREDERE_DEL25_VILLAGRES.xlsx`
- `ATUALIZACAO_DELCREDERE_DEL30_VILLAGRES.xlsx`
- `RELATORIO_AUDITORIA_VILLAGRES_DELCREDERE.json`

Resultado da auditoria:

- 271 referências lidas na fonte;
- 247 referências Villagres em cada faixa DEL;
- 24 referências Villa Art excluídas do fluxo;
- nenhuma referência Villagres duplicada;
- nenhum preço de fábrica inválido nas seis faixas;
- cálculo do preço de venda validado em todos os itens:
  `preço de fábrica exibido × 1,0536`, arredondado para duas casas;
- cada planilha possui oito colunas do layout resumido, 247 linhas, nenhuma
  fórmula e arquivo XLSX íntegro.

### Verificação Villa Art — 11/08/2026

Villa Art deve permanecer com o próprio nome; é uma linha exclusiva de alto
padrão da Villagres, não um fornecedor a ser renomeado.

Foi feito cruzamento somente de leitura entre a tabela atual, a tabela anterior
do fornecedor (`TABELA DELCREDERE VAREJO SP 2026 (1).xlsx`) e o SQLite:

- 24 referências Villa Art estão na tabela atual;
- nenhuma das 24 existe em `R3IntegradorDb.db`;
- portanto, nenhuma delas é atualização de item já importado no banco atual;
- não há referência Villa Art nova na tabela atual em relação à tabela anterior;
- a referência `123052` existia na tabela anterior e não consta mais na tabela
  atual.

Não foi criada importação nem alterada qualquer planilha nesse cruzamento.

## 11. Separação por banco: atualização e inclusão Del Credere — 11/08/2026

Foi criado o fluxo `VillagresDelcredereAtualizacaoService`, acessível no console
pela opção `17` ou pelo comando:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --atualizar-delcredere-com-banco <tabela.xlsx> R3IntegradorDb.db Saida
```

Ele consulta somente a origem `VAREJO` do SQLite e **não altera o banco**.

Resultado para a tabela `100 TABELA DEL CREDERE SP ATUALIZADA.xlsx`:

- 271 referências na fonte;
- 222 referências já cadastradas: enviadas às seis planilhas reduzidas de
  atualização;
- 49 referências ausentes no banco: enviadas a seis planilhas completas de
  importação, uma por faixa DEL;
- 25 referências do banco não aparecem na tabela atual e, portanto, não
  receberam preço nesta execução.

Para os 49 itens novos, o modelo fiscal foi obtido dos produtos cadastrados do
grupo `PORCELANATO`, NCM `69072100`. Foram herdados, entre outros, NCM, origem,
IPI, ICMS, MVA, CST, CSOSN, CFOP e os novos campos tributários. Resultado
conferido: NCM `69072100`, CSOSN `500`, CST `010`, CFOP dentro `5405` e CFOP
fora `6404` nos 49 itens.

Marca no layout de atualização/importação deve sempre refletir a marca e a
faixa Del Credere, tal como nos arquivos históricos:

- `VILLAGRES 5`, `VILLAGRES 10`, `VILLAGRES 15`, `VILLAGRES 20`,
  `VILLAGRES 25` ou `VILLAGRES 30`;
- para a linha exclusiva: `VILLA ART 5`, `VILLA ART 10`, `VILLA ART 15`,
  `VILLA ART 20`, `VILLA ART 25` ou `VILLA ART 30`.

Em 11/08/2026 foi confirmado que esse formato já existia nas saídas anteriores
Villagres. O fluxo por banco deve preservar esse padrão, tanto nas 222
atualizações quanto nas 49 inclusões (25 Villagres e 24 Villa Art).

Pasta definitiva:

- `Saida/ATUALIZACAO_E_INCLUSAO_DELCREDERE_VILLAGRES_20260811_205111/`

Conteúdo:

- seis arquivos reduzidos de atualização `VILLAGRES`, com 222 linhas cada;
  nenhum arquivo de atualização Villa Art é gerado enquanto não houver
  referência Villa Art já cadastrada no banco;
- doze arquivos completos de importação, separados por marca e faixa: seis
  `VILLAGRES` com 25 linhas e seis `VILLA_ART` com 24 linhas;
- `RELATORIO_AUDITORIA_ATUALIZACAO_E_INCLUSAO.json`.

Em todas as seis faixas, não há preço inválido e o cálculo Del Credere foi
validado como `preço de fábrica × 1,0536`, arredondado para duas casas.

Regra permanente de saída: **não gerar planilha vazia**. Para cada combinação
de marca e faixa Del Credere, gerar somente o arquivo que tiver registros:

- se houver cadastrados, gerar atualização reduzida;
- se houver novos, gerar importação completa;
- se uma das listas estiver vazia, omitir seu arquivo.

### Divergência entre ERP e SQLite — 12/08/2026

O SQLite não é a única fonte de verdade para saber se um cadastro já existe.
Uma tela de retorno do ERP confirmou que as referências abaixo já estavam no
ERP, embora ainda não constassem em `R3IntegradorDb.db`:

- `828001` a `828007`;
- `910033`, `910034`;
- `920067` a `920070`;
- `108082` a `108088`;
- `120001`, `120003`, `120004`, `120005`, `120006`.

Total: 25 referências Villagres. Para o fluxo Del Credere, essas referências
devem ir para **atualização reduzida**, nunca para importação. A fonte do ERP
prevalece para essa decisão; o banco permanece sem alteração automática, pois
não foi recebido um export/cadastro completo do ERP para sincronizá-lo.

Assim, para esta tabela, as atualizações Villagres totalizam 247 referências
(222 encontradas no SQLite + 25 confirmadas no ERP).

Em 12/08/2026, a tela de retorno do ERP ao tentar importar
`IMPORTACAO_ERP_DELCREDERE_DEL5_VILLA_ART_NOVOS.xlsx` informou: “Não existem
produtos novos para importar”. Portanto, as 24 referências Villa Art também já
existem no ERP. Como a referência é um cadastro global (não depende da faixa
DEL), elas devem ser atualizadas nas seis faixas DEL5 a DEL30 e não devem gerar
qualquer planilha de importação Villa Art.

Pasta definitiva após as confirmações do ERP:

- `Saida/ATUALIZACAO_E_INCLUSAO_DELCREDERE_VILLAGRES_20260812_134840/`

Ela contém doze atualizações: seis Villagres com 247 referências e seis Villa
Art com 24 referências. Não há planilha de importação para essas duas marcas,
pois não existem itens novos no ERP. Todos os 12 XLSX foram validados como
íntegros.

### Sincronização da réplica SQLite — 12/08/2026

Por solicitação do usuário, as 49 referências confirmadas pelo ERP e ausentes
na réplica local foram inseridas em `Produtos` com `TabelaOrigem = VAREJO`, sem
sobrescrever os cadastros existentes. A carga utilizou os dados da faixa DEL5
como preço-base na réplica e herdou o fiscal do grupo `PORCELANATO`, NCM
`69072100`.

Resultado após a sincronização:

- 296 referências VAREJO únicas no SQLite;
- 272 com marca-base `VILLAGRES`;
- 24 com marca-base `VILLA ART`;
- exemplos validados: `200029`, `828001` e `120001` com NCM `69072100`, CSOSN
  `500`, CST `010`, CFOP `5405/6404`.

Backup anterior à gravação:

- `Saida/R3IntegradorDb_ANTES_SINCRONIZACAO_ERP_20260812_135117.db.bak`

Comando adicionado para futuras sincronizações confirmadas pelo ERP:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --sincronizar-confirmados-erp <tabela.xlsx> R3IntegradorDb.db Saida
```

QA adicional da pasta definitiva: as seis atualizações Villagres possuem
exatamente a marca `VILLAGRES <faixa>` em suas 222 linhas; cada importação
Villagres possui 25 itens `VILLAGRES <faixa>` e cada importação Villa Art possui
24 itens `VILLA ART <faixa>`. Todos os 24 arquivos XLSX foram verificados como
íntegros.
