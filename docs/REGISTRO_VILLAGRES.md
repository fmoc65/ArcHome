# Registro de Processamento - Fornecedor Villagres

> Gerado em: 25/06/2026
> Contexto: Ajuste de campos fiscais com base em notas fiscais reais

---

## Sumário

1. [Problema Original](#1-problema-original)
2. [Dados das Notas Fiscais de Referência](#2-dados-das-notas-fiscais-de-referência)
3. [Alterações no Código Fonte](#3-alterações-no-código-fonte)
4. [Planilhas Geradas](#4-planilhas-geradas)
5. [Regras de Negócio Consolidadas](#5-regras-de-negócio-consolidadas)
6. [Validação dos Campos](#6-validação-dos-campos)

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
| `AliqIcmsInterna` | `18.00m` (SP destino, mantido) |
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
"PastaSaida": "/home/fernando/Projetos/Work/ARCHOME/ArcHome/R3Integrador/Saida"
```

---

## 4. Planilhas Geradas

### 4.1. Tabela Varejo (VAREJO)
- **Arquivo:** `IMPORTACAO_ERP_VAREJO_20260625_192407.xlsx`
- **Registros:** 247 produtos
- **Grupo:** PORCELANATO
- **Precificação:** Revenda (markup 75% + 10,51% impostos + frete R$ 1,50)

### 4.2. Tabela Vinílico (VINILICO)
- **Arquivo:** `IMPORTACAO_ERP_VINILICO_20260625_192408.xlsx`
- **Registros:** 16 produtos
- **NCM:** 39181000
- **SubGrupo:** CLICADO (prefixo SPC) / COLADO (prefixo LVT)

### 4.3. Tabela Lastra (LASTRA)
- **Arquivo:** `IMPORTACAO_ERP_LASTRA_20260625_192408.xlsx`
- **Registros:** 5 produtos
- **Grupo:** PORCELANATO

### 4.4. Tabelas Delcredere (12 arquivos)
- **Formato:** `IMPORTACAO_ERP_DELCREDERE_{DEL5..30}_{MARCA}_20260625_*.xlsx`
- **Marcas:** VILLAGRES (~252 registros cada) e VILLA ART (~25 registros cada)
- **Precificação:** `PrecoVenda = PrecoTabela × 1,0065` (apenas IPI 0,65%)

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
| **Delcredere** | `Venda = PrecoTabela × 1,0065` | Apenas IPI 0,65% |
| **Villa Art** | Preço tabelado do fornecedor | Sem markup, copiar coluna 20 |

### Campos Fiscais

| Campo | Revenda (Porcelanato) | Revenda (Vinílico) | Delcredere |
|---|---|---|---|
| CST | 010 | 010 | 010 |
| IPI% | 0,65 (NCM 69072100) / 0 (NCM 69072200) | 0 | 0,65 |
| AliqIcmsOrigem | 12% | 12% | 12% |
| AliqIcmsInterna | 18% | 18% | 18% |
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
| 20 | AliqIcmsInterna | `18` | ✅ |
| 26 | CST | `010` | ✅ |
| 34 | PesoBruto | `21.1` | ✅ |

### Verificado na LASTRA (ref 120003):

| Campo | Valor | Status |
|---|---|---|
| NCM | `69072200` | ✅ (diferenciado) |
| IPI% | `0` | ✅ (sem IPI p/ 69072200) |

### Verificado no VINILICO:

| Campo | Valor | Status |
|---|---|---|
| NCM | `39181000` | ✅ (antes era inválido "0") |
| SubGrupo SPC | `CLICADO` | ✅ |
| SubGrupo LVT | `COLADO` | ✅ |
| IPI% | `0` | ✅ (vinílico sem IPI) |

---

## Arquivos Modificados

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
