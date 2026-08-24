# R3 Integrador

O projeto principal da solução é o **R3Integrador.Web**, a interface web para importação e acompanhamento de atualizações de tabelas de preço.

```bash
dotnet run --project src/R3Integrador.Web/R3Integrador.Web.csproj
```

Os projetos `Application`, `Infrastructure` e `Core` dão suporte à aplicação web. O `R3Integrador.Console` permanece na solução apenas para compatibilidade com o fluxo legado.

## Banco SQLite

Para criar ou atualizar o `R3IntegradorDb.db` a partir de uma planilha no layout de
importacao do ERP:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --criar-banco /caminho/IMPORTACAO_ERP.xlsx R3IntegradorDb.db VAREJO
```

O banco possui a tabela `Produtos`, com as 60 colunas da planilha e a coluna
`TabelaOrigem`, e a tabela `Usuario` (`Nome`, `Login`, `Senha` e `Ativo`). O campo
`Ativo` aceita `0`/`1` e tem `1` (`true`) como valor padrao. Reimportar a mesma
origem substitui somente os produtos daquela origem.

Para atualizar somente o preco de fabrica dos produtos VAREJO e separar os
cadastros novos:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --atualizar-preco-fabrica /caminho/TABELA_VAREJO.xlsx R3IntegradorDb.db Saida
```

O comando usa exclusivamente a coluna `PRECO / TABELA` da planilha de origem,
mantem os demais campos dos produtos existentes e gera uma segunda planilha com
os novos produtos e os campos fiscais pendentes destacados. Para produtos cuja
unidade de venda e `M2`, a quantidade de embalagem de venda e gravada como `0`;
o valor `m2/cx` da tabela do fornecedor permanece apenas como dado logistico.

No fluxo `COM DEL CREDERE`, o preco de venda recebe sobre o preco-base os
acrescimos de `0,65%` de IPI e `4,71%` da taxa de cartao, totalizando `5,36%`.

## Adama

A Adama e tratada como representacao: `PrecoVenda` e `PrecoFabrica` recebem
diretamente a coluna `PRECO 2026` do fornecedor, sem margem, acrescimo ou
conversao para metro quadrado. A coluna `VALOR M2` nao e usada na precificacao.

Para preservar os campos fiscais da planilha validada pelo contador, corrigir
os precos, criar o backup do banco e substituir somente a origem `ADAMA`:

```bash
dotnet run --project src/R3Integrador.Console/R3Integrador.Console.csproj -- \
  --processar-adama-com-banco /caminho/TABELA_ADAMA.xlsx \
  /caminho/IMPORTACAO_ERP_ADAMA_CONTADOR_OK.xlsx \
  R3IntegradorDb.db Saida
```

Como a importacao inicial da Adama ja foi realizada, o processamento gera
`ATUALIZACAO_ADAMA_PRECOS.xlsx` no layout resumido de oito colunas do ERP, com
somente os produtos que precisam de correcao de preco. A base completa gerada
no mesmo diretorio serve exclusivamente para auditoria e sincronizacao da
replica SQLite; ela nao deve ser enviada novamente como importacao.

A opcao `16 - Atualizacao Del Credere Villagres (layout resumido)` do console gera arquivos
com apenas: marca, referencia, preco de venda, preco de fabrica, descricao,
unidade fabril, modelo e cor. Sao gerados somente os seis arquivos Villagres,
de DEL5 a DEL30, em uma pasta exclusiva por execucao. A Villa Art e o layout
completo de importacao do ERP nao fazem parte desse fluxo.

## Pipeline VillaCol

O pipeline Python audita as seis planilhas Del Credere aprovadas, extrai o PDF
de revenda, expande as referencias por embalagem, gera os arquivos corrigidos
e persiste as tabelas no SQLite:

```bash
python3 -m venv .venv-villacol
.venv-villacol/bin/pip install -r requirements-villacol.txt
.venv-villacol/bin/python scripts/villacol_pipeline.py \
  --input-dir /home/fernando/Projetos/Work/ARCHOME/Planilhas/VillaCol \
  --database R3IntegradorDb.db
```

Os arquivos de origem nao sao sobrescritos. O banco e copiado antes da carga,
e o relatorio JSON confirma as contagens do PDF, das planilhas e do SQLite.
A importacao de revenda contem todos os 63 SKUs do PDF; cada faixa Del Credere
permanece restrita aos 10 itens agregados que ja existiam na planilha aprovada.

## Auditoria Atlas

O pipeline abaixo confronta a tabela de preços Atlas com a importação aprovada,
zera a quantidade de embalagem de venda dos itens em `M2`, mantém `1` para os
itens em `PC`, remove observações e substitui somente a origem
`ATLAS_REVENDA_35` no SQLite:

```bash
.venv-villacol/bin/python scripts/atlas_audit_import.py
```
