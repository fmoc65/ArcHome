# Importacao ERP - Regras homologadas

Este documento registra as decisoes consolidadas na primeira leva de tratamento de dados
(Villagres, Villa Art, Rubinettos e Kromma). Ele deve orientar a proxima leva de
tabelas de fornecedores da Arc Home.

## Arquitetura de arquivos

- O ERP nao aceita importacao de multiplas marcas na mesma planilha.
- Cada arquivo de importacao deve conter uma unica marca.
- Tabelas de origem com marcas agrupadas, como `Villagres e Villa Art`, devem ser
  separadas em arquivos independentes.

## Referencias duplicadas

- Fornecedores sem repeticao de referencia para produtos diferentes, como Kromma e
  Villagres, podem ser importados diretamente.
- Fornecedores com a mesma referencia para produtos diferentes, como Rubinettos,
  exigem cadastro previo da marca no ERP.
- Para esses casos, a chave de atualizacao no ERP deve ser `Referencia + Descricao`.

## Limpeza de campos

- Campos como `VOLTAGEM` nao devem receber `N/A`.
- Quando nao houver valor aplicavel, o campo deve ficar vazio/nulo.
- O campo `EAN` deve preservar o codigo de barras original do fornecedor.
- Quando o fornecedor nao informar `EAN`, o campo deve ficar vazio.

## Unidades e estoque

- A unidade padrao no layout do sistema deve ser `M2`.
- A embalagem de venda deve receber a metragem minima da caixa do produto (`m2/cx`).
- Essa regra faz o ERP sugerir o arredondamento comercial correto para caixas fechadas
  no momento do orcamento.
- A embalagem de compra deve ser sempre `1,00`.
- Essa regra evita multiplicacao indevida na entrada de notas de compra emitidas pelo
  fornecedor em `m2`.

## Custo e precificacao

- O `Preco Desconto` da tabela de origem deve ser tratado como custo inicial de fabrica.
- Para revenda, o custo final deve ser:

```text
Preco fabrica com desconto + IPI + ST + R$ 1,50 de frete fixo por m2
```

- O ICMS-ST destacado na nota compoe o custo, pois e valor pago nao recuperavel.
- Para representacao/delcredere, o preco base e o preco da tabela, somando apenas o IPI.
- Exemplo homologado: Villagres usa IPI de `0,65%`.
- Como regra geral, o preco de venda e calculado por markup sobre o custo final:
  margem de `75%` mais impostos de `10,51%`.
- Quando o fornecedor trouxer preco final sugerido/tabelado, como nas linhas Villa Art,
  o valor deve ser copiado diretamente da ultima coluna da tabela de origem, sem aplicar
  markup.

## Delcredere

- Produtos representados com percentuais de comissao variados, como `DEL20`, `DEL25` e
  `DEL30`, devem ser separados em planilhas totalmente distintas por percentual.
- A separacao por percentual elimina linhas duplicadas do mesmo item dentro do mesmo
  arquivo de importacao.

## Categorizacao comercial

- Pisos, porcelanatos e ceramicas:
  - `Grupo`: `Produto`.
  - `Subgrupo`: superficie do produto, como `Natural`, `Polido` ou equivalente.
  - `Modelo`: medidas/tamanhos.
- Pisos vinilicos:
  - `Grupo`: `Vinilico`.
  - `Subgrupo`: definido obrigatoriamente pelo prefixo do codigo de fabrica.
  - Codigos iniciados por `SPC` devem usar subgrupo `CLICADO`.
  - Codigos iniciados por `LVT` devem usar subgrupo `COLADO`.
- Metais:
  - `Subgrupo`: classificado por caracteristicas estruturais, como `Bica Alta`,
    `Bica Baixa` ou `Monocomando`.
- Loucas:
  - `Subgrupo`: classificado pelo tipo de peca, como `Bacia`, `Cuba` ou `Kit Bacia`.
  - A classificacao deve usar a descricao tecnica do fabricante como base, como nas
    tabelas Roca.

## Regras fiscais

- A parametrizacao dos campos `ENQUADRAMENTO IPI`, `ALIQUOTA IBS`, `ALIQUOTA CBS` e
  `CLASSIFICACAO TRIBUTARIA` e obrigatoria.
- Esses campos devem ser validados por `NCM` e pelo estado de origem do fornecedor,
  conforme orientacao da Contabilidade Moreira.
- A aliquota interna de ICMS (`ALIQICMSINTERNA`) deve ser sempre preenchida com a
  aliquota de destino do estado da Arc Home: `SP`.

### Villagres - evidencias fiscais

PDFs verificados em `Downloads`:

- `AUTORIZACAO 1 2.pdf`:
  - Nota fiscal Villagres de porcelanato.
  - NCM observado: `69072100`.
  - Aliquota de ICMS: `12%`.
  - Aliquota de IPI: `0,65%`.
  - Valor de IPI destacado na nota.
- `WhatsApp Scan 2026-05-12 at 11.37.57.pdf`:
  - Nota fiscal Villagres da referencia `120003`.
  - NCM observado: `69072200`.
  - Aliquota de ICMS: `12%`.
  - Valor de IPI zerado/sem destaque na linha.
- `WhatsApp Scan 2026-05-12 at 12.02.10.pdf`:
  - Nota fiscal Villagres de porcelanato.
  - NCM observado: `69072100`.
  - Aliquota de ICMS: `12%`.
  - Aliquota de IPI: `0,65%`.
  - Valor de IPI destacado na nota.

### Villagres - aliquotas informadas pelo cliente

- Vinilicos:
  - `IPI %`: `0%`.
  - `ALIQICMSORIGEM`: `12%`.
  - `ALIQICMSINTERNA`: `12%`.
  - `IVA`: `66%`.
  - `PERCENTUAL ST`: `7,92%`.
  - `ALIQUOTA PIS ORIGEM`: `0,65%`.
  - `ALIQUOTA COFINS ORIGEM`: `3%`.
- Porcelanatos:
  - `IPI %`: `0,65%`.
  - `ALIQICMSORIGEM`: `12%`.
  - `ALIQICMSINTERNA`: `12%`.
  - `IVA`: `81%`.
  - `PERCENTUAL ST`: `9,86%`.
  - `ALIQUOTA PIS ORIGEM`: `0,65%`.
  - `ALIQUOTA COFINS ORIGEM`: `3,00%`.
- Ponto pendente de validacao:
  - A informacao posterior `IPI/PIS/COFINS - nao tem` conflita com as aliquotas acima
    e com duas notas fiscais de porcelanato que mostram IPI `0,65%`.
  - Antes de aplicar em massa, confirmar se `nao tem` se refere apenas a algum grupo
    especifico, como a referencia `120003`/NCM `69072200`, ou se substitui a regra geral.

## Checklist para novas tabelas

1. Identificar marcas presentes na origem e separar uma planilha por marca.
2. Verificar se ha referencias repetidas para produtos diferentes.
3. Confirmar a chave de atualizacao adequada no ERP.
4. Normalizar unidade como `M2`.
5. Preencher embalagem de venda com `m2/cx`.
6. Fixar embalagem de compra em `1,00`.
7. Preservar `EAN` original ou deixar vazio.
8. Deixar campos nao aplicaveis vazios, sem `N/A`.
9. Classificar o tipo de precificacao: revenda, representacao/delcredere ou preco final
   tabelado.
10. Separar arquivos de delcredere por percentual de comissao.
11. Deduzir `Grupo`, `Subgrupo` e `Modelo` conforme a hierarquia comercial homologada.
12. Validar campos fiscais obrigatorios por `NCM` e estado de origem.
13. Preencher `ALIQICMSINTERNA` com a aliquota de destino de SP.
