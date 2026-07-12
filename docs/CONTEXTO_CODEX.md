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

## Processamentos de 11/07/2026 - DROP, Imersi e Rubinettos

Nesta rodada foram analisadas as planilhas em
`/home/fernando/Projetos/Work/ARCHOME/Planilhas` e, inicialmente, as copias que ainda
estavam em `/home/fernando/Documentos/ArcHome`. Os arquivos de importacao foram salvos
em `R3Integrador/Saida`. Os arquivos de origem permaneceram intactos.

Convencao visual adotada nas planilhas desta rodada:

- Celulas fiscais com fundo verde: valor comprovado diretamente em XML de NF-e.
- Celulas fiscais com fundo vermelho: valor ausente, provisorio ou que depende de
  validacao/preenchimento pelo contador.
- Dados fiscais existentes apenas na tabela comercial foram preservados, mas nao foram
  tratados como evidencia equivalente a uma NF-e.

### DROP - primeira planilha pendente

Fonte comercial inicialmente usada:

- `DROP_TABELA_DE_PREÇOS_FEVEREIRO_2026_COMPLETA.xlsx`.
- Aba `LP`, linhas de produto a partir da linha 4.
- A tabela possui referencia, descricao, cor, EAN, precos, dimensoes, pesos e uma
  estimativa de ST para SP.
- A propria origem informa que a ST e estimada e deve ser conferida com a contabilidade.

Foi gerado o arquivo preliminar:

- `IMPORTACAO_ERP_DROP_PENDENTE_TRIBUTACAO_20260711.xlsx`.
- 71 produtos, marca unica `DROP`, 60 colunas e nenhum preco de venda zerado.
- Os campos fiscais sem evidencia foram deixados em vermelho.

### DROP - regeneracao com NF-e

Fontes fiscais encontradas em `/home/fernando/Projetos/Work/ARCHOME/Planilhas/Drop`:

- `NFe-739.xml`.
- `NFe-755.xml`.

Arquivo final desta etapa:

- `IMPORTACAO_ERP_DROP_COM_DADOS_NFE_20260711.xlsx`.
- 71 produtos, 60 colunas, marca unica `DROP` e nenhum preco zerado.
- As referencias comprovadas pelas notas foram `DP8205`, `DP4201` e `DP4720P`.
- Dados comprovados foram preenchidos em verde; os demais campos tributarios pendentes
  permaneceram em vermelho.

Tributacao comprovada nas NF-e DROP:

| Referencia | NCM | CEST | Origem + CST ICMS | ICMS interestadual | IPI | CST IPI |
|---|---|---|---|---:|---:|---|
| `DP8205` | `84818019` | `1007900` | `100` (`origem 1` + `CST 00`) | `4%` | `0%` | `51` |
| `DP4201` | `84818019` | `1007900` | `100` (`origem 1` + `CST 00`) | `4%` | `0%` | `51` |
| `DP4720P` | `74182000` | `1006700` | `100` (`origem 1` + `CST 00`) | `4%` | `6,5%` | `50` |

Valores comuns comprovados para os tres itens:

- Enquadramento IPI `999`.
- CST PIS `01` e aliquota PIS origem `0,65%`.
- CST COFINS `01` e aliquota COFINS origem `3%`.
- IBS `0,1%`, CBS `0,9%` e classificacao tributaria `000001`.
- Emitente `DROP METAIS`, UF `SC`, regime normal (`CRT 3`).

Pendencias DROP que nao podem ser fechadas apenas pelas notas de compra:

- `ALIQICMSINTERNA` de SP.
- IVA/MVA e percentual de ST definitivo.
- CFOPs de venda dentro e fora do estado.
- CSOSN, reducoes de base e codigo de beneficio quando aplicaveis.

### Imersi - copia para validacao contabil

Fonte disponivel:

- `IMERSI_ERP_ATE_3X_20260613_142323.xlsx`, que ja estava no layout ERP e nao era uma
  tabela comercial de origem.

Arquivo de trabalho gerado:

- `IMPORTACAO_ERP_IMERSI_ATE_3X_PENDENTE_TRIBUTACAO_20260711.xlsx`.
- 74 produtos, 60 colunas, marca unica `IMERSI` e nenhum preco zerado.
- NCM, CEST/MVA conhecidos e dados comerciais foram preservados.
- Parametros fiscais provisórios ou ainda nao homologados foram destacados em vermelho.

### Rubinettos/Kromma

Fonte comercial:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/Rubinetto2025.xlsx`.
- Aba `SP- Geral Abril 2025`.
- Cabecalho na linha 4 e produtos a partir da linha 5.
- 11.946 linhas validas e 11.892 referencias de NF distintas.
- Existem 54 referencias repetidas; elas foram mantidas porque correspondem a produtos
  ou variacoes diferentes. A chave recomendada no ERP continua sendo
  `Referencia + Descricao`.
- Nenhum produto possuia preco sugerido zerado.

A tabela comercial ja fornece, por linha:

- Referencia de NF, descricao, cor, linha, marca, EAN, unidade, dimensoes e pesos.
- Preco sugerido e diferentes precos comerciais.
- Codigo de origem, IPI, ICMS interestadual, MVA, ICMS interno, percentual de ST, NCM e
  CEST.

NF-e usadas, localizadas em
`/home/fernando/Projetos/Work/ARCHOME/Planilhas/Rubinettos`:

- `NFe-63938.xml`.
- `NFe-64439.xml` e sua copia `NFe-64439 (1).xml`.
- `NFe-64440.xml`.
- `NFe-65663.xml` e sua copia `NFe-65663 (1).xml`.
- As copias repetidas foram deduplicadas logicamente por numero da nota e referencia.

As notas comprovam 13 referencias distintas. Campos comprovados diretamente foram
marcados em verde. Regras observadas:

- Emitente `RUBINETTOS EXCLUSIVE DESIGN`, UF `RS`, `CRT 3`.
- A maior parte dos itens possui origem `5`, CST ICMS `10`, formando codigo completo
  `510`, ICMS interestadual `12%` e CFOP de entrada `6401`.
- A referencia `RB2027[4473]` possui origem `5`, CST `00`, codigo completo `500`, ICMS
  `12%` e CFOP de entrada `6107`.
- PIS: CST `01`, aliquota `0,65%`.
- COFINS: CST `01`, aliquota `3%`.
- IBS `0,1%`, CBS `0,9%` e classificacao tributaria `000001`.
- Enquadramento IPI `999` foi aproveitado quando presente nas notas.
- A maior parte dos metais das notas usa NCM `84818019`, CEST `1007900` e IPI nao
  tributado/CST `51`.
- `CUBACE21` usa NCM `73239300`, CEST `1005900`, IPI `6,5%` e CST IPI `50`.
- `CUB2001` usa NCM `73241000` e CEST `1006000`.

Marcas encontradas na origem antes da normalizacao:

- `RUBINETTOS`: 7.828 linhas.
- `RUBINETTOS COMERCIO`: 417 linhas, consolidadas como `RUBINETTOS`.
- `KROMMA`: 3.430 linhas.
- `RUBINETTOS/KROMMA`: 271 linhas cuja origem nao define uma marca unica.

Foram gerados tres arquivos porque o ERP nao aceita varias marcas no mesmo arquivo:

- `IMPORTACAO_ERP_RUBINETTOS_RUBINETTOS_COM_DADOS_NFE_20260711.xlsx`:
  8.245 produtos e 10 referencias comprovadas pelas NF-e.
- `IMPORTACAO_ERP_RUBINETTOS_KROMMA_COM_DADOS_NFE_20260711.xlsx`:
  3.430 produtos e 1 referencia comprovada pelas NF-e.
- `IMPORTACAO_ERP_RUBINETTOS_RUBINETTOS_KROMMA_COM_DADOS_NFE_20260711.xlsx`:
  271 produtos e 2 referencias comprovadas pelas NF-e.

Todos os arquivos possuem layout ERP de 60 colunas e passaram na verificacao de
integridade do formato XLSX. Os dados comerciais e a tributacao existente na tabela
foram preservados. Nos produtos presentes nas NF-e, os dados comprovados substituem ou
confirmam os valores da origem e aparecem em verde.

Pendencias Rubinettos/Kromma:

- Definir a marca individual dos 271 itens classificados como `RUBINETTOS/KROMMA` antes
  da importacao definitiva. O arquivo conjunto e apenas uma separacao fiel ao valor da
  origem.
- Validar CST, CFOPs de venda, PIS/COFINS, IBS/CBS e classificacao para os produtos que
  nao aparecem nas NF-e.
- Nao reutilizar automaticamente o CFOP de entrada das notas como CFOP de venda do ERP.
- Confirmar com o contador os campos em vermelho, incluindo CSOSN, reducoes de base,
  retencoes e codigo de beneficio quando aplicaveis.

## Tratamento Studio Morandin - 12/07/2026

Nova origem analisada:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TABELA STUDIO MORANDIN_ABR_26.xlsx`.
- Aba unica `Plan1`, com cabecalho na linha 3, produtos a partir da linha 5 e observacao
  fiscal no rodape.
- 91 produtos, 91 referencias distintas, nenhuma duplicidade e nenhum preco de revenda
  vazio ou zerado.
- 88 produtos possuem unidade `m²`; 3 produtos possuem unidade `PÇ` e foram preservados
  como `PC`, sem conversao forcada para `M2`.
- O NCM `6907.23.00` aparece apenas na observacao geral da linha 107 e foi aplicado aos
  91 produtos como informacao declarada pela origem.

Arquivo tratado gerado, ainda nao considerado importacao ERP final:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TRATADA_STUDIO_MORANDIN_ABR_26_20260712.xlsx`.
- Aba `PRODUTOS_TRATADOS`: 91 linhas e 30 colunas normalizadas.
- Aba `CRITERIOS_E_PENDENCIAS`: explica as premissas, riscos e proximos passos.
- O preco de revenda foi convertido para numero, mas ainda e necessario confirmar se ele
  representa custo de fabrica ou preco final de venda para o ERP.
- A coluna de linha comercial foi propagada para as linhas em que a origem usava celulas
  vazias como continuacao visual.
- Metragem por caixa, pecas por m², peso, pecas por caixa, medida e espessura foram
  convertidos em campos estruturados.

Tratamento fiscal adotado:

- NCM normalizado: `69072300`.
- Descricao: ladrilhos/placas de ceramica para pavimentacao ou revestimento, com absorcao
  de agua superior a 10%.
- CEST `1003000` e sujeicao a ST em SP foram classificados como regras condicionais:
  aplicam-se se os produtos forem efetivamente ladrilhos ou placas exclusivamente para
  pavimentacao/revestimento.
- IVA-ST de SP `81%` foi registrado como premissa baseada no item 24 da Portaria SRE
  88/2025, sujeito a validacao na operacao concreta.
- ICMS interno `18%`, IPI `0%` e CFOP SP `5405` foram mantidos em amarelo como premissas
  do cenario informado, nao como dados comprovados pela planilha ou NF-e.
- CST IPI `53` nao foi aplicado automaticamente: aliquota IPI zero e saida nao tributada
  sao enquadramentos diferentes.
- Origem `0 Nacional` nao foi aplicada, pois origem da mercadoria nao pode ser inferida
  pelo NCM.
- CST ICMS, CSOSN, percentual efetivo de ST, origem, CST IPI, PIS e COFINS ficaram em
  vermelho para confirmacao do contador.

Convencao visual:

- Amarelo: regra condicional ao cenario tributario informado.
- Vermelho: campo ausente ou sem comprovacao, dependente de contador, XML ou NF-e.

Proximos passos antes da importacao ERP:

1. Confirmar formalmente o regime tributario da Arc Home.
2. Confirmar com o fornecedor a origem das mercadorias.
3. Validar com o contador CST IPI, CST/CSOSN de ICMS, PIS/COFINS e percentual efetivo
   de ST para permitir o cadastro inicial sem NF-e.
4. Definir se `PRECO REVENDA` e custo de fabrica ou preco final e qual formula comercial
   deve ser aplicada.
5. Somente depois gerar o arquivo final no layout ERP de 60 colunas.
6. Revisar a parametrizacao fiscal quando for recebida a primeira NF-e real de compra.

### Modelo comercial Studio Morandin

Esclarecimento recebido em 12/07/2026:

- A Arc Home ainda nao adquiriu produtos Studio Morandin e, por isso, nao existe XML de
  NF-e de entrada disponivel neste momento.
- O fluxo comercial e de venda antecipada: primeiro ocorre a venda ao cliente e depois a
  Arc Home realiza a aquisicao correspondente junto ao fornecedor.
- A operacao trabalha com estoque minimo, sem formacao relevante de estoque previo.
- A ausencia atual de XML nao bloqueia o tratamento comercial da tabela, mas impede usar
  uma nota real como evidencia da tributacao de entrada.
- O cadastro inicial deve ser tratado como parametrizacao fiscal provisoria homologada
  pelo contador, com revisao obrigatoria apos a primeira compra e recebimento do XML.
- No ERP, `ESTOQUE MINIMO` nao deve ser preenchido automaticamente com uma quantidade
  positiva sem definicao do cliente. Usar `0` ou deixar vazio ate a politica operacional
  ser confirmada.

### Reader Studio Morandin no R3Integrador

Implementado em 12/07/2026 para tornar o processamento reproduzivel:

- Interface `IStudioMorandinReader` em `R3Integrador.Application/Interfaces`.
- Implementacao `StudioMorandinReaderService` em
  `R3Integrador.Infrastructure/Repositories`.
- Reader registrado no container de injecao, no `ImportacaoReaderSet` e no
  `ImportacaoService`.
- Opcao `9 - Processar Planilha STUDIO MORANDIN (PROVISORIA)` adicionada ao menu.
- O fluxo gera arquivos com nome
  `IMPORTACAO_ERP_STUDIO_MORANDIN_PROVISORIA_yyyyMMdd_HHmmss.xlsx` para impedir que a
  saida atual seja confundida com importacao definitiva.

Comportamento atual do reader:

- Le a aba `Plan1` a partir da linha 5.
- Propaga a linha comercial quando a origem usa celula vazia como continuacao.
- Normaliza unidade `m²` como `M2` e preserva itens vendidos por peca como `PC`.
- Usa metragem por caixa como embalagem de venda dos produtos em `M2`.
- Le peso da caixa, medida, espessura, pecas por m², pecas por caixa e preco de revenda.
- Preserva `PRECO REVENDA` provisoriamente em `PRECO VENDA` e `PRECO FABRICA`, sem
  markup, ate a regra comercial ser confirmada.
- Usa NCM `69072300`, IPI `0`, ICMS interno `18` e IVA `81` como premissas atuais.
- Mantem origem, CST/CSOSN, CFOP, percentual efetivo de ST, PIS/COFINS e IBS/CBS vazios
  ate homologacao.
- Estoque minimo e maximo ficam em `0`, coerentes com o modelo de venda antecipada e
  aquisicao posterior.

Validacao executada:

- Build completo de `R3Integrador.slnx` com `-m:1`: sem warnings e sem erros.
- Fluxo real executado com a tabela Studio Morandin: 91 produtos processados.
- Arquivo de teste gerado:
  `IMPORTACAO_ERP_STUDIO_MORANDIN_PROVISORIA_20260712_133559.xlsx`.
- O arquivo provisorio nao deve ser importado antes da homologacao fiscal e comercial.

### Regra global de cores fiscais

Definida em 12/07/2026 para todas as novas planilhas e readers:

- Verde (`#C6EFCE`): dado fiscal confirmado por fonte confiavel, como XML/NF-e, ou
  declarado expressamente na tabela de origem.
- Amarelo (`#FFEB9C`): estimativa, premissa ou regra condicional ainda sujeita a
  homologacao.
- Vermelho (`#FFC7CE`): campo fiscal ausente ou pendente de contador/evidencia.

A regra foi incorporada ao codigo:

- `ProdutoErpDto` possui o metadado `SituacaoCamposFiscais`, sem alterar as 60 colunas
  posicionais do ERP.
- `ExcelExportService` aplica automaticamente as cores declaradas por cada reader.
- Novos readers devem classificar explicitamente os campos como `Confirmado`, `Estimado`
  ou `Pendente`.
- O reader Studio Morandin foi o primeiro a usar a regra global: NCM verde; IPI, ICMS
  interno e IVA amarelos; demais campos fiscais sem evidencia em vermelho.
- Arquivo validado com a nova formatacao:
  `IMPORTACAO_ERP_STUDIO_MORANDIN_PROVISORIA_20260712_134103.xlsx`, com 91 produtos e
  60 colunas.

## Tratamento Invita - 12/07/2026

Origem analisada:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TabeladePrecosINVITACNPJICMS-18% ABR2026.xlsx`.
- O nome correto da marca e `INVITA`; a solicitacao inicial mencionava Invicta.
- Aba `CNPJ`, cabecalho na linha 91 e produtos a partir da linha 92.
- Tabela vigente originalmente de 15/04/2026 a 10/05/2026; portanto esta vencida na data
  do tratamento.
- A origem declara fornecedor nacional, centro de distribuicao em Paulinia/SP, ICMS de
  18% e precos com IPI, ICMS e substituicao tributaria inclusos.
- 49 linhas de produto, 48 produtos unicos. A linha `KITFILTRO`, EAN
  `7898593939646`, aparece duas vezes e foi deduplicada por `Codigo + EAN`.
- Todos os 48 produtos possuem EAN e precos positivos.

NCMs encontrados:

- `84146000`: 9 coifas.
- `85167100`: 2 cafeteiras.
- `73211100`: 5 churrasqueiras/cooktops/fornos a gas.
- `85166000`: 6 cooktops/fornos/fogao eletricos.
- `85165000`: 1 micro-ondas; nao constava no levantamento fiscal inicial.
- `84181000`: 4 refrigeradores.
- `84186999`: 17 adegas, beer centers e maquinas de gelo.
- `73239300`: 1 espatula de aluminio para pizza.
- `84212100`: 1 kit filtro de agua apos deduplicacao.
- `84501200`: 1 lava e seca.
- `84221100`: 1 lava-loucas; nao constava no levantamento fiscal inicial.

Ponto critico de classificacao:

- O item `I-PA-FPZ-12-XX-NMHA` esta descrito como `ESPATULA EM ALUMINIO`, mas usa NCM
  `73239300`, cuja descricao se refere a artefatos de aco inoxidavel. Exigir confirmacao
  do fornecedor/contador antes da importacao.

Reader implementado:

- Interface `IInvitaReader`.
- Implementacao `InvitaReaderService`.
- Registrado no `ImportacaoReaderSet`, injecao de dependencia e `ImportacaoService`.
- Opcao `10 - Processar Planilha INVITA (PROVISORIA)` adicionada ao menu.
- Le precos numericos diretamente das celulas para preservar valores com separador de
  milhar.
- Usa `Preco Sugerido com Desconto` como preco de venda provisório e `Preco Lojista com
  Impostos` como preco de fabrica/custo provisório.
- Preserva NCM, EAN, familia, descricao, voltagem, desconto, IPI, ICMS, PIS e COFINS da
  origem.
- Estoques minimo/maximo ficam zerados.

Arquivo gerado e validado:

- `IMPORTACAO_ERP_INVITA_PROVISORIA_20260712_135746.xlsx`.
- 48 produtos, 60 colunas, estrutura XLSX integra.
- Verde: NCM, IPI, ICMS origem/interno e PIS/COFINS origem declarados por produto na
  tabela.
- Vermelho: IVA, CST/CSOSN, CFOP, percentual ST, enquadramento IPI, IBS/CBS,
  classificacao tributaria e codigo de beneficio.

Alerta temporal de ICMS-ST em SP:

- A Portaria SRE 34/2026 foi publicada em 30/06/2026, mas entra em vigor apenas em
  `01/10/2026`.
- Em 12/07/2026 nao e correto tratar automaticamente refrigeradores e demais produtos
  atingidos como ja excluidos da ST.
- A situacao deve ser parametrizada conforme a data da operacao: regra vigente ate
  30/09/2026 e nova regra a partir de 01/10/2026.
- A noticia do Governo de SP confirma a exclusao futura, nao uma exclusao ja vigente.

Necessario para o contador fechar a Invita:

1. Confirmar o regime tributario da Arc Home e se devem ser usados CST ou CSOSN.
2. Definir, por NCM + descricao, a sujeicao a ST ate 30/09/2026 e a partir de 01/10/2026.
3. Informar CEST, IVA/MVA e percentual efetivo de ST dos itens ainda sujeitos.
4. Definir CST/CSOSN e CFOP de venda dentro/fora de SP para itens com e sem ST.
5. Confirmar CST IPI e enquadramento IPI para as diferentes aliquotas da origem.
6. Homologar PIS/COFINS da tabela (`1,65%`/`7,60%`) para o cadastro da Arc Home.
7. Informar IBS, CBS, classificacao tributaria e codigo de beneficio.
8. Corrigir ou confirmar o NCM `73239300` da espatula descrita como aluminio.
9. Validar especificamente o kit filtro `84212100` pela descricao e finalidade, evitando
   enquadrar ST somente pelo NCM.
10. Confirmar se os precos usados pelo reader correspondem corretamente a custo e venda
    no modelo comercial da Arc Home.

## Tratamento Derosso - 12/07/2026

Origem:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TABELADEPREÇOSDEROSSO_REPRESENTAÇÃO_01_05.xlsx`.
- Aba `DEROSSO_REPRESENTAÇÃO_01.05.26`.
- Tabela de precos praticados para representacao comissionada, vigente a partir de
  01/05/2026.
- 169 produtos, 169 SKUs distintos e nenhum preco zerado.
- Unidades: 90 produtos em `M2`, 76 em `PC` e 3 em `KG`.
- Todos os produtos possuem NCM `69041000` declarado pela origem.
- A origem nao fornece EAN, UF/origem da mercadoria, IPI, ICMS, PIS/COFINS, CST/CSOSN,
  CFOP, CEST, IVA/MVA, IBS/CBS ou classificacao tributaria.

Dados comerciais aproveitados:

- Produto, cor, SKU, unidade, preco unitario, preco por peca, dimensoes, embalagem,
  quantidade por embalagem, peso, dimensoes da embalagem e consumo por m².
- Para produtos vendidos em `M2`, a embalagem de venda e calculada como
  `1 / quantidade de embalagens por m²`.
- Para produtos em `PC` ou `KG`, a embalagem de venda fica em `1`.
- Por ser tabela de representacao comissionada, o preco unitario foi preservado
  provisoriamente em `PRECO VENDA` e `PRECO FABRICA`, sem markup.

Reader implementado:

- Interface `IDerossoReader`.
- Implementacao `DerossoReaderService`.
- Registrado no `ImportacaoReaderSet`, injecao de dependencia e `ImportacaoService`.
- Opcao `11 - Processar Planilha DEROSSO (PROVISORIA)` adicionada ao menu.

Tratamento fiscal e cores:

- NCM `69041000`: verde, pois consta em todas as linhas da origem.
- ICMS interno `12%`: amarelo. A legislacao paulista preve tratamento de 12% para
  tijolos ceramicos nao esmaltados nem vitrificados, mas e necessario confirmar que
  todos os produtos Derosso atendem a essa descricao.
- Percentual ST `0`: amarelo. O item da Portaria CAT 68/2019 para posicao `6904`, CEST
  historico `1002700`, foi revogado a partir de 01/01/2026 pela Portaria SRE 64/2025.
- IPI, ICMS origem, IVA, CST/CSOSN, CFOP, enquadramento IPI, PIS/COFINS, IBS/CBS,
  classificacao tributaria e beneficio: vermelho.
- O NCM unico precisa ser confirmado especialmente para garrafeiras, suportes,
  elementos 3D e seixos soltos; nao se deve concluir a classificacao apenas por serem
  ceramicos.

Arquivo gerado e validado:

- `IMPORTACAO_ERP_DEROSSO_PROVISORIA_20260712_141208.xlsx`.
- 169 produtos e 60 colunas.
- Integridade XLSX validada e build completo sem warnings/erros.

Necessario para o contador/fornecedor fechar a Derosso:

1. Confirmar NCM `69041000` por familia, principalmente garrafeiras, suportes,
   elementos 3D e seixos.
2. Confirmar que os produtos sao ceramicos nao esmaltados/nem vitrificados para aplicar
   ICMS interno de 12%.
3. Confirmar formalmente a ausencia de ST em SP e se algum item recebe outro CEST pela
   descricao/finalidade.
4. Informar UF e origem fiscal da mercadoria.
5. Informar IPI, CST IPI e enquadramento IPI.
6. Definir CST/CSOSN e CFOP dentro/fora de SP para o regime da Arc Home.
7. Informar PIS/COFINS de origem aplicaveis.
8. Informar IBS, CBS, classificacao tributaria e codigo de beneficio.
9. Confirmar a regra comercial da representacao: se o preco unitario deve ser copiado
   diretamente para venda e fabrica ou se existe comissao/fator adicional.

### Derosso Revenda

Nova origem recebida em 12/07/2026:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TABELA_DE_PRECOS_DEROSSO_REVENDA_01_05.xlsx`.
- Aba `DEROSSO_REVENDA_01.05.26`.
- 142 produtos e SKUs unicos, sem precos zerados.
- 78 produtos em `M2` e 64 em `PC`.
- Todos usam NCM `69041000` na origem.
- O catalogo de revenda e menor que o de representacao: nao contem os tres itens em
  `KG` e possui 27 produtos a menos.
- Os precos sao diferentes da representacao e nao podem ser misturados.

O `DerossoReaderService` foi adaptado para identificar automaticamente `REVENDA` ou
`REPRESENTACAO` pelo nome da aba/titulo. Os arquivos futuros passam a incluir o tipo:

- `IMPORTACAO_ERP_DEROSSO_REVENDA_PROVISORIA_*.xlsx`.
- `IMPORTACAO_ERP_DEROSSO_REPRESENTACAO_PROVISORIA_*.xlsx`.

Arquivo de revenda gerado e validado:

- `IMPORTACAO_ERP_DEROSSO_REVENDA_PROVISORIA_20260712_141938.xlsx`.
- 142 produtos e 60 colunas, sem erro de integridade.
- NCM verde; ICMS interno `12%` e ST `0%` amarelos; demais campos fiscais pendentes em
  vermelho.
- O preco unitario da origem foi preservado provisoriamente em `PRECO FABRICA` e
  `PRECO VENDA`. A planilha nao deve ser importada ate a margem/markup da revenda ser
  definida.

Pendencia comercial adicional para a revenda:

- Confirmar se `PRECO UNITARIO DE VENDA` representa o custo de aquisicao da Arc Home.
- Definir markup/margem, frete e demais componentes para calcular o preco final ao
  consumidor.
- Nao aplicar automaticamente a regra comercial da representacao aos itens de revenda.

## Tratamento Atlas Revenda 35% - 12/07/2026

Origem:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TABELA ATLAS REVENDA 35% - MAIO 2026.xlsx`.
- Arquivo exportado do Numbers com quatro abas: `Resumo da Exportação`, `Plan1`, `Plan2`
  e `Plan3`.
- `Resumo da Exportação` contem apenas metadados da conversao.
- `Plan2` e `Plan3` estao realmente vazias.
- Todos os produtos e regras fiscais estao em `Plan1`.

Estrutura identificada:

- 194 produtos regulares com referencia e EAN.
- 9 registros adicionais de cantoneiras/cantos sem EAN, usando o tipo como referencia.
- Total exportado: 203 produtos, todos com preco positivo.
- 150 referencias textuais distintas nos produtos regulares; referencias repetidas
  representam formatos/categorias diferentes e possuem EANs distintos.
- Manter chave ERP `Referencia + Descricao`; a descricao gerada inclui secao, cor e
  formato para diferenciar as variacoes.

Regras fiscais declaradas por secao na origem:

- NCM `69073000`: 110 produtos; IPI `0,65%`; ST informada como isenta/zero.
- NCM `69072200`: 72 produtos; IPI `0,65%`; percentual ST `9,86%`.
- NCM `69072100`: 12 produtos; IPI `0,65%`; percentual ST `9,86%`.
- NCM `69074000`: 9 cantoneiras/cantos; IPI `0,65%`; ST informada como isenta/zero.
- Para NCMs `69072100` e `69072200`, IVA `81%` foi incluído em amarelo como estimativa
  baseada na regra paulista de revestimentos ceramicos.
- Para `69073000` e `69074000`, IVA ficou pendente/vermelho; ST isenta nao foi tratada
  automaticamente como IVA cadastral zero.

Reader implementado:

- Interface `IAtlasReader`.
- Implementacao `AtlasReaderService`.
- Le apenas `Plan1` e ignora de forma segura as abas vazias/metadados.
- Herda NCM, IPI e ST do cabecalho fiscal ativo de cada bloco.
- Trata o layout regular e o layout alternativo de cantoneiras/cantos.
- Registrado no `ImportacaoReaderSet`, injecao de dependencia e `ImportacaoService`.
- Opcao `12 - Processar Planilha ATLAS REVENDA 35% (PROVISORIA)` adicionada ao menu.

Arquivo gerado e validado:

- `IMPORTACAO_ERP_ATLAS_REVENDA_35_PROVISORIA_20260712_143227.xlsx`.
- 203 produtos, 60 colunas, marca unica `ATLAS`, nenhum preco zerado e integridade XLSX
  validada.
- Verde: NCM, IPI e percentual ST informados pela origem.
- Amarelo: IVA `81%` estimado nos grupos com ST `9,86%`.
- Vermelho: ICMS origem/interno, CST/CSOSN, CFOP, enquadramento IPI, PIS/COFINS,
  IBS/CBS, classificacao tributaria e demais pendencias.

Necessario para o contador/cliente fechar a Atlas:

1. Confirmar CEST por NCM e descricao: especialmente diferenciar revestimentos,
   pastilhas/mosaicos e cantoneiras/cantos.
2. Homologar se `69073000` e `69074000` estao realmente sem ST em SP em cada descricao.
3. Confirmar se `9,86%` e o percentual efetivo de ST esperado pelo ERP e homologar IVA
   `81%` para `69072100`/`69072200`.
4. Informar UF/origem da mercadoria e ICMS origem/interno.
5. Definir CST/CSOSN e CFOP dentro/fora de SP.
6. Confirmar CST IPI e enquadramento IPI para a aliquota `0,65%`.
7. Informar PIS/COFINS de origem.
8. Informar IBS, CBS, classificacao tributaria e codigo de beneficio.
9. Confirmar codigos/EANs e desdobramento por cor dos 9 itens de cantoneiras/cantos que
   aparecem de forma agregada na origem.
10. Confirmar se os precos publicados ja incorporam o desconto de 35% e se representam
    custo da Arc Home ou preco final; definir markup/frete antes da importacao definitiva.

## Conversao Nina Martinelli - 12/07/2026

Arquivo recebido:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/TABELA-PRECO-NINA-MARTINELLI_2026-Representação-REV02.numbers`.
- `.numbers` e o formato proprietario de planilha do Apple Numbers.
- O arquivo possui 14 MB, cerca de 200 blocos internos de tabelas `.iwa` e mais de mil
  imagens de produtos; por isso aplicativos sem suporte ao Numbers podem exibir apenas
  as imagens/previas.
- LibreOffice nao conseguiu converter o arquivo diretamente.

Foi usado temporariamente o pacote `numbers-parser` para ler os blocos estruturados do
Numbers. O pacote foi instalado somente em `/tmp`, sem dependencia adicionada ao
R3Integrador.

Folhas encontradas:

- `Indice`: sem produtos relevantes.
- `Coleção Completa`: 1.141 linhas ativas e 1.138 codigos distintos; quatro bordas
  especiais compartilham o texto `depende do raio` como codigo.
- `Valores`: planilha auxiliar com formulas/erros de calculo.
- `Lançamentos`: 224 produtos, todos ja contidos em `Coleção Completa`; nao devem ser
  duplicados.
- `Politica Comercial`: regras de representacao.
- `Descontinuados`: 68 produtos, excluidos da conversao ativa.

Distribuicao da colecao ativa:

- 736 produtos em `m²`.
- 396 produtos em `pç`.
- 9 produtos em `L (litro)`.
- Todos informam UF `SP` e possuem preco de representacao positivo.
- A origem nao possui EAN ou campos fiscais/NCM visiveis na tabela principal analisada.

Arquivo XLSX estruturado gerado:

- `/home/fernando/Projetos/Work/ARCHOME/Planilhas/CONVERTIDA_NINA_MARTINELLI_2026_REPRESENTACAO_REV02.xlsx`.
- Aba `COLECAO_COMPLETA`, com 1.141 produtos e os 22 campos comerciais recuperados.
- Imagens nao foram incorporadas na copia estruturada; os dados comerciais foram
  priorizados para permitir tratamento e criacao futura do reader.

Proximos passos:

1. Definir referencias tecnicas unicas para as quatro bordas com codigo `depende do raio`.
2. Mapear grupo/subgrupo e normalizar unidades/embalagens.
3. Obter NCM e tributacao por familia com fornecedor/contador.
4. Criar `NinaMartinelliReaderService` sobre o XLSX convertido.
5. Gerar planilha ERP provisoria com campos fiscais em vermelho.
