# Contexto Codex - ArcHome / R3Integrador

> Atualizado em: 01/07/2026
> Objetivo: reconstruir o contexto de trabalho apos migracao do Windows para Linux Mint.

## Visao geral

O projeto principal analisado e `R3Integrador`, uma solucao .NET `net10.0` para
normalizar planilhas de fornecedores e gerar arquivos Excel no layout de importacao
do ERP da Arc Home.

Estrutura atual:

- `R3Integrador.Application`: DTOs, interfaces, mapper para o layout ERP e servico de
  orquestracao.
- `R3Integrador.Core`: enums e base de dominio ainda enxuta.
- `R3Integrador.Infrastructure`: readers de planilhas por fornecedor/tabela e exportador
  Excel com ClosedXML.
- `R3Integrador.Console`: menu interativo para selecionar o fornecedor/processamento.
- `R3Integrador/Saida`: arquivos `.xlsx` ja gerados para importacao ERP.

Commits recentes relevantes:

- `626b61a` - Valores corretos de Vinilico.
- `4ef5b5f` - Colocado valores das notas fiscais.
- `d198622` - Inicio da retomada com documento de regras e ajustes Delcredere.
- Commits anteriores incluem Imersi, Roca, Rubinettos, Delcredere, Vinilico, Lastra e
  Varejo.

## Documentos existentes

- `docs/importacao-erp-regras.md`: regras homologadas gerais para novas tabelas.
- `docs/REGISTRO_VILLAGRES.md`: historico detalhado do tratamento Villagres, notas
  fiscais usadas como evidencia, planilhas geradas e correcao de Vinilico em 27/06/2026.

## Regras homologadas principais

- O ERP nao aceita multiplas marcas na mesma planilha; gerar um arquivo por marca.
- Fornecedores com referencia repetida para produtos diferentes, como Rubinettos, exigem
  chave de atualizacao `Referencia + Descricao` no ERP.
- Campos nao aplicaveis devem ficar vazios; nao usar `N/A`.
- `EAN` deve preservar o codigo original quando existir; quando nao existir, deixar vazio.
- Unidade padrao para pisos/porcelanatos/vinilicos no layout: `M2`.
- Embalagem de venda deve receber `m2/cx`.
- Embalagem de compra deve ser `1,00`.
- `ALIQICMSINTERNA` deve considerar destino SP.
- Campos `ENQUADRAMENTO IPI`, `ALIQUOTA IBS`, `ALIQUOTA CBS` e
  `CLASSIFICACAO TRIBUTARIA` sao obrigatorios e dependem de NCM/origem.

## Precificacao

Revenda Villagres/Vinilico/Lastra:

```text
Custo final = Preco Desconto * 1,1051 + 1,50
Preco venda = Custo final * 1,75
```

Delcredere:

```text
Preco venda = Preco tabela * 1,0065
```

Villa Art:

```text
Preco venda = copiar coluna final/tabelada da origem, sem markup.
```

## Villagres e Villa Art

Fluxos implementados:

- `VAREJO`: `ExcelReaderService`, aba `VAREJO`.
- `VINILICO`: `VinilicoReaderService`, aba `VINILICO`.
- `LASTRA`: `LastraReaderService`, aba `LASTRA`.
- `COM DEL CREDERE`: `DelcredereReaderService`, separa por `DEL5`, `DEL10`, `DEL15`,
  `DEL20`, `DEL25`, `DEL30` e por marca `VILLAGRES`/`VILLA ART`.
- `VILLA ART - BOUTIQUE`: `VillaArtReaderService`.

Regras fiscais consolidadas no mapper:

- Porcelanato:
  - Grupo `PORCELANATO`.
  - NCM padrao `69072100`.
  - Referencia `120003` usa NCM `69072200`.
  - IPI `0,65%` apenas para NCM `69072100`; NCM `69072200` fica com IPI `0`.
  - CST `010`.
  - ICMS origem `12%`.
  - ICMS interna `12%`.
  - IVA `81%`.
  - Percentual ST `9,86%`.
  - PIS origem `0,65`.
  - COFINS origem `3`.
  - Enquadramento IPI `999`.
  - IBS `0,1`, CBS `0,9`, classificacao tributaria `000001`.
- Vinilico:
  - Grupo `VINILICO`.
  - NCM `39181000`.
  - CST `010`.
  - IPI `0`.
  - ICMS origem `12%`.
  - ICMS interna `12%`.
  - IVA `66%`.
  - Percentual ST `7,92%`.
  - PIS origem `0,65`.
  - COFINS origem `3`.
  - Subgrupo por prefixo: `SPC` => `CLICADO`, `LVT` => `COLADO`.
- Delcredere:
  - Gerado em arquivos separados por percentual e marca.
  - Marca exportada recebe percentual, por exemplo `VILLAGRES 20`.
  - IPI forca `0,65%` e campos CBS/IBS/classificacao sao preenchidos para porcelanato.

Ponto de atencao documentado:

- Houve informacao posterior do cliente dizendo `IPI/PIS/COFINS - nao tem`, mas isso
  conflita com notas fiscais de porcelanato que mostram IPI `0,65%`. Confirmar antes
  de alterar em massa.

## Rubinettos

Fluxo: `RubinettosReaderService`.

- Le todas as abas.
- Gera um arquivo por marca.
- Grupo `METAIS`.
- Preserva codigo de barras da coluna 31.
- Usa unidade da origem, com fallback `UN`.
- Usa CST `060`, CFOP dentro `5405`, CFOP fora `6404`.
- ICMS interna usa a aliquota informada; se vazia, fallback `18%`.
- Observacao inclui aba e CEST quando existir.

## Roca / Celite

Fluxo: `RocaReaderService`.

- Le a primeira aba.
- Gera um arquivo por marca.
- Normaliza `CE` para `CELITE` e `RO` para `ROCA`.
- Grupo `LOUCAS E METAIS`.
- Preco final usa campanha quando houver; caso contrario, preco tabela.
- Preserva EAN/codigo de barras.
- Peso liquido/bruto vem da origem.
- Observacao inclui dimensoes do produto, embalagem, CEST e origem quando existirem.

## Imersi

Fluxo: `ImersiReaderService`, aba `LISTA DE PRECOS`.

- Gera arquivos separados por tabela:
  - `7X_A_10X`
  - `4X_A_6X`
  - `ATE_3X`
  - `ANTECIPADO`
- Marca `IMERSI`.
- Grupo por NCM:
  - `39221000` => `BANHEIRAS`
  - `84818019` => `METAIS`
  - demais => `BANHO`
- Tributacao conhecida:
  - `39221000`: CEST `1001300`, MVA original `91`, ajustada 4% `123,61`,
    ajustada 12% `104,98`.
  - `84818019`: CEST `1007900`, MVA original `85`, ajustada 4% `116,59`,
    ajustada 12% `98,54`.

## Layout ERP

`ProdutoErpDto` possui 60 colunas, exportadas por `ExcelExportService` em ordem posicional.

Campos relevantes:

- Coluna 11 `VOLTAGEM`: vazio.
- Coluna 13 `NCM`.
- Coluna 15 `PRECO VENDA`.
- Coluna 18 `IPI %`.
- Coluna 19 `ALIQICMSORIGEM`.
- Coluna 20 `ALIQICMSINTERNA`.
- Coluna 21 `IVA`.
- Coluna 24 `UNIDADE`.
- Coluna 25 `QTDE EMBALAGEM DE VENDA`.
- Coluna 26 `CST`.
- Coluna 34 `PESOBRUTO`.
- Coluna 39 `PERCENTUAL ST`.
- Coluna 51 `ENQUADRAMENTO IPI`.
- Coluna 52 `ALIQUOTA PIS ORIGEM`.
- Coluna 53 `ALIQUOTA COFINS ORIGEM`.
- Colunas 57-59 `ALIQUOTA IBS`, `ALIQUOTA CBS`, `CLASSIFICACAO TRIBUTARIA`.

## Arquivos ja gerados

Em `R3Integrador/Saida` existem arquivos gerados em 25/06/2026 e 27/06/2026:

- `IMPORTACAO_ERP_VAREJO_20260625_192407.xlsx`
- `IMPORTACAO_ERP_LASTRA_20260625_192408.xlsx`
- `IMPORTACAO_ERP_VINILICO_20260625_192408.xlsx`
- `IMPORTACAO_ERP_VINILICO_20260627_134234.xlsx`
- `IMPORTACAO_ERP_VINILICO_20260627_134317.xlsx`
- 12 arquivos `IMPORTACAO_ERP_DELCREDERE_*` separados por `DEL5..DEL30` e marca.

Arquivos finais gerados e validados em 01/07/2026 a partir das planilhas em
`/home/fernando/Documentos/ArcHome`:

- Fonte `TABELAVAREJO2026ARCHOME .xlsx`:
  - `IMPORTACAO_ERP_VAREJO_20260701_200706.xlsx`: 247 linhas.
  - `IMPORTACAO_ERP_VINILICO_20260701_200706.xlsx`: 16 linhas.
  - `IMPORTACAO_ERP_LASTRA_20260701_200706.xlsx`: 5 linhas.
- Fonte `ROCA 5101-ARC HOME MAT  P CONST LTDA.xlsx`:
  - `IMPORTACAO_ERP_ROCA_CELITE_20260701_200706.xlsx`: 10 linhas.
  - `IMPORTACAO_ERP_ROCA_ROCA_20260701_200707.xlsx`: 1427 linhas.

Validacao feita nos arquivos finais:

- Sem preco de venda zerado.
- Sem NCM vazio ou `0`.
- Varejo: 247 itens com NCM `69072100`, IPI `0,65`, ICMS interna `12`, IVA `81`,
  ST `9,86`, PIS origem `0,65`, COFINS origem `3`.
- Vinilico: 16 itens com NCM `39181000`, IPI `0`, ICMS interna `12`, IVA `66`,
  ST `7,92`, PIS origem `0,65`, COFINS origem `3`.
- Lastra: 4 itens com NCM `69072100` e IPI `0,65`; referencia `120003` com NCM
  `69072200` e IPI `0`; todos com ICMS interna `12`, IVA `81`, ST `9,86`,
  PIS origem `0,65`, COFINS origem `3`.
- ROCA/CELITE e ROCA/ROCA: separados por marca, NCM/preco presentes, CST `060`,
  ICMS interna `18`, IPI/IVA/ST vindos da planilha de origem.

Planilhas em `/home/fernando/Documentos/ArcHome` que nao foram geradas nesta rodada:

- `DROP_TABELA_DE_PREÇOS_FEVEREIRO_2026_COMPLETA.xlsx`: parece ser DROP/Rubinettos,
  mas o layout nao bate com o reader atual de Rubinettos; gerar agora produziria
  colunas erradas.
- `IMERSI_ERP_ATE_3X_20260613_142323.xlsx`: ja e uma planilha de saida ERP, nao fonte.
- `IMPORTACAO_ERP_DELCREDERE_DEL20_20260609_200256.xlsx`: ja e uma planilha de saida
  ERP, nao fonte.

## Pendencias fiscais e combinados com cliente

Registro de 01/07/2026 para continuidade nas proximas interacoes com o cliente.

Planilhas com informacao fiscal suficiente na propria origem:

- `ROCA 5101-ARC HOME MAT  P CONST LTDA.xlsx`:
  - Traz `Classificacao Fiscal`/NCM, `% Ipi`, `% Icms`, `% St`, `Valor St`, `CEST`
    e `Origem`.
  - A importacao ROCA deve aproveitar esses campos diretamente da planilha.
  - Mesmo assim, quando houver NF de entrada, confrontar CST/CFOP/PIS/COFINS e confirmar
    se a tributacao da entrada real bate com o que foi gerado.

Planilhas sem informacao fiscal completa na origem:

- `TABELAVAREJO2026ARCHOME .xlsx`:
  - Abas `VAREJO`, `VINILICO` e `LASTRA` nao trazem NCM, IPI, ICMS, IVA/MVA, ST,
    PIS, COFINS, CEST ou CST.
  - O sistema preenche esses campos por regra do mapper.
  - Vinilico foi confirmado pelo cliente em 01/07/2026:
    - IPI `0%`.
    - ICMS origem `12%`.
    - ICMS saida/interna `12%`.
    - IVA `66%`.
    - ST `7,92%`.
    - PIS origem `0,65%`.

## Atualizacao 10/07/2026

- Resolvido o problema de execucao interativa do console no Linux: o menu agora aceita sequencia de stdin para selecionar opcao, fornecer caminho e sair.
- Gerado com sucesso o arquivo Roca/Celite a partir de `/home/fernando/Documentos/ArcHome/ROCA 5101-ARC HOME MAT  P CONST LTDA.xlsx`:
  - `Saida/IMPORTACAO_ERP_ROCA_CELITE_20260710_202405.xlsx`
  - `Saida/IMPORTACAO_ERP_ROCA_ROCA_20260710_202405.xlsx`
- Confirmado que o fluxo Roca gera campos fiscais e de tributos no layout ERP, incluindo `IPI %`, `ALIQICMSORIGEM`, `ALIQICMSINTERNA`, `IVA`, `PERCENTUAL ST`, `CST`, `CFOP`, `CEST`, `PESO LIQUIDO` e `PESO BRUTO`.
- Identificado que o arquivo `/home/fernando/Documentos/ArcHome/IMERSI_ERP_ATE_3X_20260613_142323.xlsx` nao e fonte bruta de Imersi, mas sim planilha ja no formato de importacao ERP.
- Ajustado o `ImersiReaderService` para considerar a aba alternativa `IMPORTACAO_ERP` quando `LISTA DE PREÇOS` nao existe.
- Conclusao: para gerar Imersi, precisamos da planilha de origem Imersi raw; o arquivo atual nao tem os campos de tabela de preco esperados e, portanto, nao produz registros.
    - COFINS origem `3%`.
  - Varejo/Lastra de porcelanato estao preenchidos por regra atual:
    - NCM padrao `69072100`, exceto referencia `120003` com NCM `69072200`.
    - IPI `0,65%` para `69072100` e `0%` para `69072200`.
    - ICMS origem `12%`, ICMS interna `12%`, IVA `81%`, ST `9,86%`,
      PIS origem `0,65%`, COFINS origem `3%`.
    - Confrontar com NF de entrada quando o cliente enviar.

- `DROP_TABELA_DE_PREÇOS_FEVEREIRO_2026_COMPLETA.xlsx`:
  - Tem referencia, descricao, cor, EAN/codigo de barras, preco, peso e uma coluna de
    estimativa de substituicao tributaria para SP.
  - Nao traz NCM, CEST, CST, IPI, ICMS, PIS/COFINS completos no layout analisado.
  - Nao gerar importacao final sem NF de entrada, XML, tabela fiscal do fornecedor ou
    mapeamento fiscal validado por NCM/produto.

Uso de NF de entrada:

- O combinado com o cliente e que, ao receber notas fiscais de entrada, elas sejam usadas
  para confrontar a tributacao aplicada pelo integrador.
- Preferir XML da NF-e ao PDF/DANFE, porque o XML traz os impostos por item de forma
  estruturada.
- Com XML de NF-e e possivel validar por item:
  - NCM.
  - CEST.
  - CFOP.
  - CST/CSOSN.
  - Origem da mercadoria.
  - IPI.
  - ICMS.
  - ICMS-ST.
  - Base de calculo e valores de imposto.
  - PIS.
  - COFINS.
- A NF de entrada e a melhor evidencia para calibrar as regras por fornecedor/NCM. Alguns
  campos de venda no ERP ainda podem depender de regra fiscal por UF/NCM, mas a NF deve
  ser usada para detectar divergencias e ajustar o mapper antes de novas geracoes.

## Pauta para cliente/contador - 11/07/2026

Objetivo da reuniao: fechar as regras fiscais pendentes antes de tratar as planilhas
como finais para importacao ERP ou envio ao contador.

Pontos conceituais para alinhar:

- `CST` nao permite descobrir `NCM`; ele indica a situacao tributaria da operacao.
- `NCM` ajuda a definir tributacao, mas sozinho nao fecha todos os campos.
- Para substituicao tributaria, a validacao segura depende de `NCM + CEST + descricao
  do produto + UF/origem/destino + regime da empresa`.
- Para vendas da loja de Sorocaba para Sorocaba/regiao, usar regra interna de SP pode
  ser uma premissa comercial/fiscal, mas deve ser confirmada pelo contador.
- `PIS` e `COFINS` nao dependem de Sorocaba; dependem do regime da empresa/produto.
  Se a Arc Home estiver em regra cumulativa, `PIS 0,65%` e `COFINS 3%` podem ser
  plausiveis. Se for regra nao cumulativa, os percentuais podem ser outros.

Perguntas objetivas para levar ao cliente/contador:

1. Confirmar o regime tributario usado para preencher `ALIQUOTA PIS ORIGEM` e
   `ALIQUOTA COFINS ORIGEM` nas importacoes do ERP.
2. Confirmar se, para mercadorias com `CST 060` e venda dentro de SP, o ERP deve receber:
   `CFOP DENTRO = 5405`, `CFOP FORA = 6404`, `ALIQICMSINTERNA = 18` e campos de ST
   conforme a tabela do fornecedor.
3. Confirmar se Roca/Celite deve receber `PIS ORIGEM = 0,65` e `COFINS ORIGEM = 3`,
   ou se esses campos devem ser preenchidos por outra regra.
4. Confirmar se os campos da reforma tributaria no layout ERP (`ALIQUOTA IBS`,
   `ALIQUOTA CBS`, `CLASSIFICACAO TRIBUTARIA`) podem ficar `0` para Roca/Celite e
   fornecedores de metais/loucas, ou se o contador exige codificacao especifica.
5. Pedir XML de NF-e de entrada, preferencialmente uma ou mais notas por fornecedor:
   Roca/Celite, Drop/Rubinettos, Imersi e demais fornecedores pendentes.
6. Pedir tabela fiscal do fornecedor quando houver: `NCM`, `CEST`, `CST/CSOSN`,
   `IPI`, `ICMS`, `MVA/IVA`, `PIS`, `COFINS`, origem e observacoes por produto.
7. Confirmar se a chave de atualizacao no ERP para fornecedores com referencias
   repetidas deve continuar sendo `Referencia + Descricao`.

Roca/Celite:

- Os arquivos de 10/07/2026 estao estruturalmente corretos: marca separada, NCM
  presente, EAN presente, preco presente, `CST 060`, `CFOP 5405/6404` e ICMS interna
  `18`.
- Pendencia: `ALIQUOTA PIS ORIGEM` e `ALIQUOTA COFINS ORIGEM` estao vazias em todas
  as linhas. Nao enviar como fiscalmente completo sem confirmar a regra.
- Decisao a pedir: preencher `0,65` e `3` por padrao, deixar vazio, ou preencher outro
  percentual conforme regime/contabilidade.

Drop/Rubinettos:

- A planilha `DROP_TABELA_DE_PREÇOS_FEVEREIRO_2026_COMPLETA.xlsx` tem referencia,
  descricao, cor, codigo de barras, preco, peso/dimensoes e uma coluna `SP` com
  estimativa de substituicao tributaria.
- A propria planilha informa que a ST e estimativa e deve ser conferida com a
  contabilidade.
- Na planilha analisada nao foram encontrados `NCM`, `CEST`, `CST`, `CFOP`, `IPI`,
  `PIS` ou `COFINS`.
- Nao gerar importacao fiscal final de Drop sem uma destas evidencias:
  - XML/NF-e de entrada com impostos por item.
  - Tabela fiscal do fornecedor com `NCM` e `CEST`.
  - Mapeamento validado pelo contador por referencia/produto.

Delcredere/Villagres/Villa Art:

- Arquivos de 10/07/2026 foram validados: layout de 60 colunas, marca unica por arquivo,
  precos positivos, NCM preenchido, formula `PrecoVenda = PrecoFabrica * 1,0065`,
  `CST 010`, ICMS origem/interna `12`, IVA `81`, ST `9,86`, PIS `0,65`, COFINS `3`.
- A referencia `120003` saiu corretamente com NCM `69072200` e `IPI 0`.
- Pedir apenas confirmacao final se a regra `PIS 0,65` e `COFINS 3` tambem esta
  homologada para estes arquivos de representacao/delcredere.

## Ambiente Linux Mint

Estado em 01/07/2026:

- SDK instalado: .NET `10.0.109`.
- Runtime: .NET `10.0.9`.
- Target dos projetos: `net10.0`.
- `git status` estava limpo antes da criacao deste documento.
- O caminho de log foi ajustado de `C:\\log\\log-.txt` para `logs/log-.txt`.
- As dependencias `Dapper`, `Microsoft.Data.Sqlite` e pacotes Serilog foram removidas
  de `R3Integrador.Infrastructure`, pois nao havia uso direto no codigo atual. O pacote
  SQLite trazia alerta NuGet `NU1903` via dependencia transitiva.
- A referencia direta de `R3Integrador.Infrastructure` para `R3Integrador.Core` foi
  removida porque era redundante; `Infrastructure` usa os contratos de `Application`,
  e `Application` ja referencia `Core`.
- O `Console` referencia explicitamente `Application` e `Infrastructure`, pois usa tipos
  das duas camadas.
- `DisableTransitiveProjectReferences` foi habilitado nos projetos para impedir que o
  SDK adicione referencias transitivas sem metadata de framework.
- Os `ProjectReference` restantes receberam `SkipGetTargetFrameworkProperties="true"`
  e `SetTargetFramework="TargetFramework=net10.0"`, pois todos os projetos sao
  single-target `net10.0` e o MSBuild 18 no Linux estava falhando silenciosamente ao
  executar `GetTargetFrameworks` em referencias de projeto.
- `BuildInParallel` foi desativado nos projetos para evitar disputa de arquivos quando
  `Console` e `Infrastructure` constroem a cadeia `Application/Core` no mesmo build.
- Para build da solucao `.slnx`, usar `-m:1`, porque o agendamento da solucao ainda pode
  paralelizar projetos apesar das propriedades dos `.csproj`.
- O parser decimal foi centralizado em `DecimalParser.Parse`, porque algumas planilhas
  trazem valores com ponto decimal (`66.4627`, `1.42`) e outras podem trazer formato
  brasileiro; remover pontos incondicionalmente transformava `66.4627` em `664627`.

Comandos testados:

```bash
dotnet build src/R3Integrador.Core/R3Integrador.Core.csproj --no-restore -v minimal
dotnet build src/R3Integrador.Application/R3Integrador.Application.csproj --no-restore -v minimal
```

Resultado: ambos compilam.

Comandos validados apos a correcao:

```bash
dotnet restore R3Integrador.slnx --disable-parallel -v minimal
dotnet build src/R3Integrador.Console/R3Integrador.Console.csproj --no-restore -v minimal
dotnet build R3Integrador.slnx --no-restore -m:1 -v minimal
```

Resultado: build concluido sem warnings e sem erros.

Falhas encontradas antes da correcao:

- `Infrastructure`/`Console` retornavam falha por auditoria NuGet `NU1903` em
  `SQLitePCLRaw.lib.e_sqlite3` 2.1.11, trazido por `Microsoft.Data.Sqlite`.
- Havia arquivos de log herdados do Windows com nome literal `C:\log\...` dentro da
  arvore do projeto, gerados pelo caminho antigo do Serilog.
- `dotnet build R3Integrador.slnx` sem `-m:1` ainda pode falhar silenciosamente no SDK
  instalado por disputa/agendamento paralelo; usar o comando validado acima.

## Como continuar

Para novas tabelas de fornecedor:

1. Identificar marcas presentes e separar uma planilha por marca.
2. Verificar referencias duplicadas.
3. Definir chave de atualizacao ERP.
4. Preservar EAN quando houver.
5. Normalizar campos vazios sem `N/A`.
6. Definir unidade e embalagem de venda/compra.
7. Definir regra de preco: revenda, delcredere/representacao ou preco final tabelado.
8. Classificar grupo/subgrupo/modelo.
9. Validar NCM, CST, IPI, ICMS, IVA, ST, PIS/COFINS, IBS/CBS e classificacao tributaria.
10. Gerar arquivo por marca/tabela e validar colunas criticas no Excel final.
