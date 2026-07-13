Assunto: Validacao contabil das planilhas ERP — processamentos de 10 a 12/07/2026

Ola, [Nome da cliente], tudo bem?

Concluimos o tratamento das tabelas recebidas entre 10 e 12 de julho e geramos os
arquivos no layout de 60 colunas do ERP. Poderia, por gentileza, encaminhar este resumo
ao contador para validacao antes de qualquer importacao definitiva?

Como tratamos os dados

Preservamos os dados comerciais disponiveis em cada origem, como codigo, descricao,
marca, EAN, unidade, preco, dimensoes e peso. Quando havia XML de NF-e, ele foi usado
como evidencia primaria por item. Quando a origem ou a NF informava NCM, usamos essa
classificacao como ponto de partida para consultar bases publicas e legislacao e montar
premissas fiscais iniciais. Isso nao significa que o NCM, sozinho, defina toda a regra:
CST/CSOSN, CEST, ST, IVA/MVA, CFOP, origem, PIS/COFINS e os novos campos IBS/CBS
dependem tambem da descricao, operacao, UF e regime tributario.

Nas planilhas, verde indica dado comprovado por XML/NF-e ou declarado na fonte; amarelo
indica premissa que precisa de homologacao; vermelho indica campo pendente.

Planilhas para validacao

1. 10/07 — ROCA e CELITE
Foram gerados dois arquivos, um por marca. NCM, EAN, preco, CST 060, CFOP 5405/6404 e
ICMS interno de 18% vieram da tabela. Precisamos confirmar a regra de PIS/COFINS,
IBS/CBS/classificacao tributaria e se os CFOPs permanecem corretos para as vendas da
Arc Home.

2. 10/07 — Delcredere Villagres e Villa Art
Foram gerados 12 arquivos, separados por percentual DEL5 a DEL30 e por marca. A formula
comercial foi aplicada e os dados de porcelanato foram estruturados. Pedimos confirmacao
de PIS 0,65%/COFINS 3%, IPI nos casos de divergencia e campos IBS/CBS/classificacao
tributaria.

3. 11/07 — DROP
Considerar a planilha `IMPORTACAO_ERP_DROP_COM_DADOS_NFE_20260711.xlsx`. Tres
referencias foram confirmadas por XML de NF-e. Para concluir todo o catalogo, precisamos
da regra de ICMS interno de SP, IVA/MVA e percentual de ST, CFOP de venda dentro/fora
de SP, CSOSN, reducoes e codigo de beneficio quando aplicaveis.

4. 11/07 — Imersi
O arquivo de origem ja estava no layout ERP; por isso foi gerada uma copia de trabalho
com 74 itens e dados comerciais preservados. Solicitamos confirmacao de NCM/CEST/MVA e
de toda a tributacao por produto antes da importacao final.

5. 11/07 — Rubinettos e Kromma
Foram gerados tres arquivos por marca. Treze referencias foram confrontadas com XML.
Precisamos validar os campos fiscais dos itens sem NF-e, definir a marca dos 271 itens
classificados como `RUBINETTOS/KROMMA` e confirmar os CFOPs de venda — os CFOPs de
entrada das notas nao foram copiados automaticamente para venda.

6. 12/07 — Studio Morandin
Arquivo provisório com 91 itens. NCM 69072300 consta como informacao geral da tabela;
IPI 0%, ICMS interno 18% e IVA 81% foram apenas premissas condicionais. Precisamos de
origem, CST/CSOSN, CFOP, ST efetiva, PIS/COFINS, enquadramento IPI e confirmacao se o
preco de revenda e custo ou preco final.

7. 12/07 — Invita
Arquivo provisório com 48 itens. NCM, EAN, IPI, ICMS, PIS e COFINS vieram da tabela,
mas precisamos definir ST por NCM+descricao e por data de operacao, CEST, IVA/MVA,
CST/CSOSN, CFOP, enquadramento IPI, IBS/CBS e confirmar o NCM da espatula de aluminio.

8. 12/07 — Derosso (representacao e revenda)
O NCM 69041000 veio da origem. ICMS 12% e ST zero foram deixados apenas como premissas.
Precisamos confirmar NCM por familia, ausencia de ST em SP, origem, IPI, CST/CSOSN,
CFOP, PIS/COFINS, IBS/CBS e a formacao do preco: especialmente margem, frete e markup
na tabela de revenda.

9. 12/07 — Atlas Revenda 35%
NCM, IPI e ST foram aproveitados da tabela; IVA 81% e somente uma estimativa para os
grupos com ST 9,86%. Precisamos de CEST, confirmacao de ST/IVA por descricao, origem,
ICMS, CST/CSOSN, CFOP, PIS/COFINS, IBS/CBS e confirmacao do preco apos o desconto de
35%.

10. 12/07 — Nina Martinelli
Arquivo provisório com 1.141 itens. Foram extraidos codigo, descricao, unidade, preco e
dados de embalagem; quatro bordas receberam referencias tecnicas unicas porque a origem
trazia o mesmo texto de codigo. A tabela nao possui NCM, EAN ou dados fiscais. Precisamos
do mapeamento de NCM e tributacao por familia, alem da confirmacao da regra comercial de
representacao.

11. 12/07 — Tabela Especial SL (PDF)
Arquivo provisório com 216 itens. O PDF forneceu codigo, descricao, linha, formato e
preco FOB Tatuí e declara que os precos estao sem IPI, ST e DIFAL. Nao ha NCM, EAN,
embalagem, peso ou tributacao por item. Precisamos do cadastro fiscal do fornecedor ou
XML de NF-e, com NCM, CEST, origem, ICMS, IPI, PIS/COFINS, CST/CSOSN, CFOP, ST/IVA e
dos parametros comerciais para transformar o preco FOB em preco de venda.

Para todos os fornecedores, pedimos tambem a confirmacao do regime tributario da Arc
Home e o preenchimento de IBS, CBS, classificacao tributaria e codigo de beneficio no
layout do ERP. Se possivel, o formato mais seguro para fechar os itens e XML de NF-e de
entrada recente, acompanhado da regra fiscal de venda para SP e para fora do estado.

Ficamos a disposicao para aplicar as respostas na proxima geracao dos arquivos.

Atenciosamente,

[Seu nome]
