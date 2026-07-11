# Registro de Processamento - Fornecedor Villagres

> Gerado em: 25/06/2026
> Contexto: Ajuste de campos fiscais com base em notas fiscais reais
>
> Atualizado em: 27/06/2026
> Contexto: Correção de alíquotas específicas para Vinílico conforme feedback do cliente (Kaption AI)
>
> Atualizado em: 10/07/2026
> Contexto: Consolidação das regras atuais do mapper e validação dos arquivos Delcredere gerados em 10/07/2026

---

## Sumário

1. [Problema Original](#1-problema-original)
2. [Dados das Notas Fiscais de Referência](#2-dados-das-notas-fiscais-de-referência)
3. [Alterações no Código Fonte](#3-alterações-no-código-fonte)
4. [Planilhas Geradas](#4-planilhas-geradas)
5. [Regras de Negócio Consolidadas](#5-regras-de-negócio-consolidadas)
6. [Validação dos Campos](#6-validação-dos-campos)
7. [Correção de Alíquotas Vinílico (27/06/2026)](#7-correção-de-alíquotas-vinílico-27062026)
8. [Validação dos Arquivos de 10/07/2026](#8-validação-dos-arquivos-de-10072026)
9. [Pontos para Fechar com Cliente/Contador](#9-pontos-para-fechar-com-clientecontador)

---

## 1. Problema Original

Cliente reportou que "Villagres já temos notas então pelo que entendi todos os campos deveriam estar preenchidos, mas ainda tem campos em branco".

**Campos que estavam em branco (incorretos):**
| Campo | Antes | Depois |
|---|---|---|
| Voltagem | `"N/A"` | `string.Empty` |
| PesoBruto | `0` | Lido da planilha (col O) |
| IPI% | `0` | `0,65%` (NCM 69072100) / `0%` (demais) |
| AliqIcmsOrigem | `18%` | `12%` |
| CST | `060` | `010` |
| NCM Vinílico | `0` (inválido) | `39181000` |

---

## 2. Dados das Notas Fiscais de Referência

### NF-e 582239 (Renascence Natural - Ref 630062A-R209)

| Campo | Valor |
|---|---|
| Produto | PORC. RENASCENCE NATURAL 631X108 |
| NCM | 69072100 |
| CST | 010 (ICMS + ST) |
| CFOP | 5401 |
| Preço Unitário | R$ 51,62/m² |
| ICMS | 12% (R$ 389,81) |
| ICMS ST | R$ 320,34 (Base R$ 5.917,90) |
| IPI | 0,65% (R$ 21,11) |
| MVA (ST) | ~82,17% |

### NF-e 582240 (Pietra Viva Natural - Ref 120003A-L120)

| Campo | Valor |
|---|---|
| Produto | PORC. PIETRA VIVA NATURAL 119,5X250 |
| NCM | 69072200 |
| CST | 010 (ICMS + ST) |
| CFOP | 5403 |
| Preço Unitário | R$ 271,87/m² |
| ICMS | 12% (R$ 682,83) |
| ICMS ST | R$ 553,09 (Base R$ 10.299,33) |
| IPI | 0% |
| MVA (ST) | ~81,00% |

### Composição de Custo (confirmada pelas NFs)

```
CustoFinal = PrecoDesconto × (1 + IPI%) + ICMS-ST/m² + Frete R$ 1,50/m²

Onde:
  IPI% = 0,65% para NCM 69072100
  IPI% = 0%    para NCM 69072200 e Vinílicos
  Frete = R$ 1,50/m² (fixo comercial, incluso no markup)
```

---

## 3. Alterações no Código Fonte

### 3.1. `ProdutoErpDto.cs` (linha 18)
```csharp
// ANTES:
public string Voltagem { get; set; } = "N/A";

// DEPOIS:
public string Voltagem { get; set; } = string.Empty;
```

### 3.2. `ProdutoNormalizado.cs` (linha 23)
Adicionada propriedade `PesoBrutoM2` para capturar peso bruto/m² da planilha.

### 3.3. `ProdutoErpMapper.cs` - Principais alterações

| Método | Comportamento |
|---|---|
| `ObterNcm()` | Vinílico → `39181000`; Porcelanato → `69072100` (exceto ref `120003` → `69072200`) |
| `ObterIpiPercentual()` | NCM `69072100` → `0,65%`; demais → `0%` |
| `ObterCst()` | Porcelanato/Vinílico → `010`; demais → `060` |
| `AliqIcmsOrigem` | `12.00m` (antes `18.00m`) |
| `AliqIcmsInterna` | Porcelanato/Vinílico → `12.00m`; demais → `18.00m` |
| `Iva` | Porcelanato → `81.00m`; Vinílico → `66.00m`; demais → `0` |
| `PercentualSt` | Porcelanato → `9.86m`; Vinílico → `7.92m`; demais → `0` |
| `AliquotaPisOrigem` | Porcelanato/Vinílico → `"0,65"`; demais → vazio |
| `AliquotaCofinsOrigem` | Porcelanato/Vinílico → `"3"`; demais → vazio |
| `PesoBruto` | `produto.PesoBrutoM2` (antes `0`) |

### 3.4. Readers (ExcelReader, VinilicoReader, LastraReader, VillaArtReader, DelcredereReader)

Todos os readers foram alterados para ler a **coluna 15 (peso bruto/m²)** da planilha.

Em cada classe `Linha*` interna, adicionado:
```csharp
public decimal PesoBrutoM2 { get; private set; }
```

E no método `Atualizar()`:
```csharp
var pesoBruto = worksheet.Cell(row, 15).GetString();
if (!string.IsNullOrWhiteSpace(pesoBruto))
    PesoBrutoM2 = ParseDecimal(pesoBruto);
```

### 3.5. `appsettings.json`
```json
"Diretorios": {
  "PastaSaida": "/home/fernando/Projetos/Work/ARCHOME/ArcHome/R3Integrador/Saida"
}
```

---

## 4. Planilhas Geradas

### 4.1. Tabela Varejo (VAREJO)
- **Arquivo:** `IMPORTACAO_ERP_VAREJO_20260625_192407.xlsx`
- **Registros:** 247 produtos
- **Grupo:** PORCELANATO
- **Precificação:** Revenda (markup 75% + 10,51% impostos + frete R$ 1,50)

### 4.2. Tabela Vinílico (VINILICO)
- **Arquivo 25/06:** `IMPORTACAO_ERP_VINILICO_20260625_192408.xlsx`
- **Arquivo 27/06 (corrigido):** `IMPORTACAO_ERP_VINILICO_20260627_134317.xlsx`
- **Registros:** 16 produtos
- **NCM:** 39181000
- **SubGrupo:** CLICADO (prefixo SPC) / COLADO (prefixo LVT)
- **Diferença:** Alíquotas fiscais ajustadas conforme dados do cliente (ver seção 7)

### 4.3. Tabela Lastra (LASTRA)
- **Arquivo:** `IMPORTACAO_ERP_LASTRA_20260625_192408.xlsx`
- **Registros:** 5 produtos
- **Grupo:** PORCELANATO

### 4.4. Tabelas Delcredere (12 arquivos)
- **Formato:** `IMPORTACAO_ERP_DELCREDERE_{DEL5..30}_{MARCA}_20260625_*.xlsx`
- **Marcas:** VILLAGRES (~252 registros cada) e VILLA ART (~25 registros cada)
- **Precificação:** `PrecoVenda = PrecoTabela × 1,0065`
- **Observação fiscal atual:** o fator comercial de `1,0065` permanece para Delcredere,
  mas o campo `IPI %` segue a regra por NCM: `0,65%` para `69072100` e `0%` para
  `69072200`.

### 4.5. Villa Art (não processado)
- **Status:** Aguardando arquivo separado do fornecedor
- **Aba esperada:** "VILLA ART - BOUTIQUE"
- **Precificação:** Preço tabelado direto (sem markup)

---

## 5. Regras de Negócio Consolidadas

### Precificação

| Tipo | Fórmula | Observação |
|---|---|---|
| **Revenda** (VAREJO/VINILICO/LASTRA) | `Custo = PrecoDesconto × 1,1051 + 1,50` | 10,51% ≈ IPI 0,65% + ST ~9,86% |
| | `Venda = Custo × 1,75` | Markup de 75% |
| **Delcredere** | `Venda = PrecoTabela × 1,0065` | Fator comercial fixo; campo IPI segue o NCM |
| **Villa Art** | Preço tabelado do fornecedor | Sem markup, copiar coluna 20 |

### Campos Fiscais

| Campo | Revenda (Porcelanato) | Revenda (Vinílico) | Delcredere |
|---|---|---|---|
| CST | 010 | 010 | 010 |
| IPI% | 0,65 (NCM 69072100) / 0 (NCM 69072200) | 0 | 0,65 (NCM 69072100) / 0 (NCM 69072200) |
| AliqIcmsOrigem | 12% | 12% | 12% |
| AliqIcmsInterna | 12% | 12% | 12% |
| IVA | 81% | 66% | 81% |
| PercentualST | 9,86% | 7,92% | 9,86% |
| AliqPisOrigem | 0,65% | 0,65% | 0,65% |
| AliqCofinsOrigem | 3% | 3% | 3% |
| CFOP Dentro | 5405 | 5405 | 5405 |
| CFOP Fora | 6404 | 6404 | 6404 |
| NCM | 69072100 / 69072200 | 39181000 | 69072100 / 69072200 |

### Hierarquia

| Tipo | Grupo | SubGrupo | Modelo |
|---|---|---|---|
| Porcelanato | PORCELANATO | Superfície (NATURAL, POLIDO, EXTERNO, etc.) | Medidas (20X141,5, 92X92, etc.) |
| Vinílico SPC | VINILICO | CLICADO | Medidas |
| Vinílico LVT | VINILICO | COLADO | Medidas |
| Lastra | PORCELANATO | Superfície | Medidas |

---

## 6. Validação dos Campos

### Verificado no VAREJO (primeiros produtos):

| Coluna | Campo | Valor | Status |
|---|---|---|---|
| 11 | Voltagem | `""` (vazio) | ✅ |
| 13 | NCM | `69072100` | ✅ |
| 15 | PrecoVenda | `131.16` | ✅ |
| 18 | IPI% | `0.65` | ✅ |
| 19 | AliqIcmsOrigem | `12` | ✅ |
| 20 | AliqIcmsInterna | `12` | ✅ |
| 21 | IVA | `81` | ✅ |
| 26 | CST | `010` | ✅ |
| 34 | PesoBruto | `21.1` | ✅ |
| 39 | PercentualST | `9.86` | ✅ |
| 52 | AliqPisOrigem | `0,65` | ✅ |
| 53 | AliqCofinsOrigem | `3` | ✅ |

### Verificado na LASTRA (ref 120003):

| Campo | Valor | Status |
|---|---|---|
| NCM | `69072200` | ✅ (diferenciado) |
| IPI% | `0` | ✅ (sem IPI p/ 69072200) |

### Verificado no VINILICO (25/06):

| Campo | Valor | Status |
|---|---|---|
| NCM | `39181000` | ✅ (antes era inválido "0") |
| SubGrupo SPC | `CLICADO` | ✅ |
| SubGrupo LVT | `COLADO` | ✅ |
| IPI% | `0` | ✅ (vinílico sem IPI) |

### Verificado no VINILICO (27/06 - após correção):

| Coluna | Campo | Valor | Status |
|---|---|---|---|
| 19 | AliqIcmsOrigem | `12` | ✅ |
| 20 | AliqIcmsInterna | `12` | ✅ (corrigido de 18) |
| 21 | IVA | `66` | ✅ (corrigido de 0) |
| 39 | PercentualST | `7,92` | ✅ (corrigido de 0) |
| 52 | AliqPisOrigem | `0,65` | ✅ (corrigido de vazio) |
| 53 | AliqCofinsOrigem | `3` | ✅ (corrigido de vazio) |

---

## 7. Correção de Alíquotas Vinílico (27/06/2026)

### Problema

Cliente (Kaption AI) reportou que a tabela de importação dos vinílicos estava com `ALIQICMSINTERNA = 18%`, mas o valor correto informado anteriormente era `12%`. Aproveitou-se para preencher também os demais campos fiscais específicos dos vinílicos que estavam zerados/vazios.

### Dados do Cliente (Vinílicos)

| Parâmetro | Valor |
|---|---|
| IPI | 0% |
| AliqIcmsOrigem | 12% |
| AliqIcmsInterna | 12% |
| IVA | 66% |
| Percentual ST | 7,92% |
| PIS Origem | 0,65% |
| COFINS Origem | 3% |

### Alterações no `ProdutoErpMapper.cs`

Os campos passaram a ser condicionais por tipo de produto:

```csharp
AliqIcmsInterna = EhPorcelanato(produto) || EhVinilico(produto) ? 12.00m : 18.00m,
Iva = EhVinilico(produto) ? 66.00m : EhPorcelanato(produto) ? 81.00m : 0,
PercentualSt = EhVinilico(produto) ? 7.92m : EhPorcelanato(produto) ? 9.86m : 0,
AliquotaPisOrigem = EhPorcelanato(produto) || EhVinilico(produto) ? "0,65" : string.Empty,
AliquotaCofinsOrigem = EhPorcelanato(produto) || EhVinilico(produto) ? "3" : string.Empty
```

### Planilha Gerada

- **Arquivo:** `IMPORTACAO_ERP_VINILICO_20260627_134317.xlsx`
- **Registros:** 16 produtos
- **Status:** Todos os campos fiscais validados conforme acima

---

## Arquivos Modificados

### 25/06/2026
```
src/R3Integrador.Application/DTOs/ProdutoErpDto.cs
src/R3Integrador.Application/DTOs/ProdutoNormalizado.cs
src/R3Integrador.Application/Mappers/ProdutoErpMapper.cs
src/R3Integrador.Infrastructure/Repositories/ExcelReaderService.cs
src/R3Integrador.Infrastructure/Repositories/VinilicoReaderService.cs
src/R3Integrador.Infrastructure/Repositories/LastraReaderService.cs
src/R3Integrador.Infrastructure/Repositories/VillaArtReaderService.cs
src/R3Integrador.Infrastructure/Repositories/DelcredereReaderService.cs
src/R3Integrador.Console/appsettings.json
```

### 27/06/2026
```
src/R3Integrador.Application/Mappers/ProdutoErpMapper.cs
```

### 10/07/2026
```
docs/REGISTRO_VILLAGRES.md
```

---

## 8. Validação dos Arquivos de 10/07/2026

Arquivos analisados em `R3Integrador/Saida`:

- 12 arquivos `IMPORTACAO_ERP_DELCREDERE_*_20260710_*.xlsx`.
- 2 arquivos `IMPORTACAO_ERP_ROCA_*_20260710_*.xlsx`.

### Delcredere

Validações realizadas:

- Todos os arquivos possuem 60 colunas no layout ERP.
- Arquivos Villagres possuem 252 linhas cada.
- Arquivos Villa Art possuem 25 linhas cada.
- Cada arquivo contém apenas uma marca, com o percentual no nome exportado:
  `VILLAGRES 5`, `VILLAGRES 10`, `VILLAGRES 15`, `VILLAGRES 20`,
  `VILLAGRES 25`, `VILLAGRES 30`, `VILLA ART 5`, `VILLA ART 10`,
  `VILLA ART 15`, `VILLA ART 20`, `VILLA ART 25`, `VILLA ART 30`.
- Todos os registros usam `Grupo = PORCELANATO`, `Unidade = M2`, `CST = 010`,
  `ALIQICMSORIGEM = 12`, `ALIQICMSINTERNA = 12`, `IVA = 81` e
  `PERCENTUAL ST = 9,86`.
- Nenhum arquivo possui preço de venda zerado.
- Nenhum arquivo possui NCM vazio ou `0`.
- A fórmula `PREÇO VENDA = PREÇO DE FÁBRICA × 1,0065`, arredondada para 2 casas,
  foi validada em todos os registros.
- Nos arquivos Villagres, a referência `120003` saiu com NCM `69072200` e `IPI % = 0`,
  conforme exceção fiscal já registrada.

Resultado: arquivos Delcredere de 10/07/2026 estão coerentes com as regras atuais.

### Roca/Celite

Validações realizadas:

- `IMPORTACAO_ERP_ROCA_CELITE_20260710_202405.xlsx`: 10 linhas, marca única `CELITE`.
- `IMPORTACAO_ERP_ROCA_ROCA_20260710_202405.xlsx`: 1427 linhas, marca única `ROCA`.
- Ambos possuem 60 colunas no layout ERP.
- Todos os registros usam `Grupo = LOUCAS E METAIS`, `Unidade = UN`,
  `CST = 060` e `ALIQICMSINTERNA = 18`.
- Nenhum registro possui preço de venda zerado.
- Nenhum registro possui NCM vazio ou `0`.
- Nenhum registro possui EAN/código de barras vazio.
- Ponto de atenção: `ALIQUOTA PIS ORIGEM` e `ALIQUOTA COFINS ORIGEM` estão vazias
  em todas as linhas dos dois arquivos Roca/Celite. Esses campos não vêm mapeados no
  reader atual e devem ser validados com o contador antes de tratar os arquivos como
  finais.

Resultado: arquivos Roca/Celite de 10/07/2026 estão estruturalmente coerentes com o
reader atual, mas não devem ser considerados fiscalmente completos sem validação de
PIS/COFINS origem.

---

## 9. Pontos para Fechar com Cliente/Contador

Levar estes pontos para a reunião de 11/07/2026 antes de tratar os arquivos como
definitivos para importação ou envio contábil.

### O que pode ser inferido e o que não pode

- `CST` não permite descobrir `NCM`. O CST indica a situação tributária da operação,
  não identifica a mercadoria.
- `NCM` ajuda a definir impostos, mas não fecha sozinho todos os campos fiscais.
- Para substituição tributária, a validação deve usar `NCM + CEST + descrição do
  produto + UF de origem/destino + regime da empresa`.
- Operações da loja de Sorocaba para Sorocaba/região podem usar premissas internas de
  SP, mas isso precisa ser confirmado pelo contador.
- `PIS` e `COFINS` dependem do regime fiscal/produto, não da cidade. Confirmar se a
  regra aplicável é `PIS 0,65%` e `COFINS 3%` ou outra.

### Roca/Celite

- Arquivos de 10/07/2026 estão estruturalmente consistentes.
- Pendência: `ALIQUOTA PIS ORIGEM` e `ALIQUOTA COFINS ORIGEM` estão vazias em todas
  as linhas.
- Perguntar ao contador se deve preencher `0,65` e `3`, deixar vazio, ou aplicar outra
  regra por NCM/produto.
- Confirmar se `ALIQUOTA IBS`, `ALIQUOTA CBS` e `CLASSIFICACAO TRIBUTARIA` podem ficar
  `0` para Roca/Celite ou se devem receber codificação específica.

### Drop/Rubinettos

- A planilha DROP analisada possui referência, descrição, cor, código de barras, preço,
  peso/dimensões e coluna `SP` com estimativa de ST.
- Não foram encontrados `NCM`, `CEST`, `CST`, `CFOP`, `IPI`, `PIS` ou `COFINS` na
  planilha DROP.
- A própria planilha informa que a ST é estimativa e deve ser conferida com a
  contabilidade.
- Pedir ao cliente/contador uma das opções:
  - XML/NF-e de entrada com impostos por item.
  - Tabela fiscal do fornecedor com `NCM`, `CEST`, `CST/CSOSN`, `IPI`, `ICMS`,
    `MVA/IVA`, `PIS` e `COFINS`.
  - Mapeamento fiscal validado por referência/produto.

### Delcredere/Villagres/Villa Art

- Os arquivos de 10/07/2026 passaram na validação de layout, marca, quantidade, NCM,
  IPI por NCM, ICMS, IVA, ST, PIS/COFINS e fórmula de preço.
- Confirmar apenas se `PIS 0,65` e `COFINS 3` estão homologados para todos os itens de
  representação/delcredere.
