using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class AdamaReaderService : IAdamaReader
{
    private const string NomeAba = "TABELA DE PREÇOS 2026";
    private const string Marca = "ADAMA";
    private const int LinhaCabecalho = 1;
    private const int PrimeiraLinhaProduto = 2;

    private static readonly string[] CabecalhosEsperados =
    [
        "CODIGO",
        "DESCRICAO DE",
        "DESCRICAO",
        "COR",
        "UN",
        "ACABAMENTO",
        "LINHA",
        "MEDIDA",
        "FATOR CONVERSAO",
        "PRECO 2026",
        "VALOR M2",
        "PESO (KG)",
        "POS.IPI/NCM"
    ];

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        using var workbook = new XLWorkbook(caminhoArquivo);
        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException(
                $"Aba '{NomeAba}' nao encontrada na tabela Adama.");
        }

        ValidarCabecalho(worksheet);

        var produtos = new List<ProdutoErpDto>();
        var assinaturasPorReferencia = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var duplicadosIdenticosIgnorados = 0;

        for (var row = PrimeiraLinhaProduto; row <= ultimaLinha; row++)
        {
            var referencia = NormalizarTexto(worksheet.Cell(row, 1).GetString());
            var descricaoCompleta = NormalizarTexto(worksheet.Cell(row, 3).GetString());

            if (string.IsNullOrWhiteSpace(referencia) && string.IsNullOrWhiteSpace(descricaoCompleta))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(referencia) || string.IsNullOrWhiteSpace(descricaoCompleta))
            {
                throw new InvalidOperationException(
                    $"Referencia ou descricao ausente na linha {row} da tabela Adama.");
            }

            var assinatura = CriarAssinatura(worksheet, row);
            if (assinaturasPorReferencia.TryGetValue(referencia, out var assinaturaExistente))
            {
                if (!assinatura.Equals(assinaturaExistente, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Referencia '{referencia}' repetida com dados divergentes na linha {row}.");
                }

                duplicadosIdenticosIgnorados++;
                continue;
            }

            var ncm = SomenteDigitos(worksheet.Cell(row, 13).GetString());
            if (ncm.Length != 8)
            {
                throw new InvalidOperationException(
                    $"NCM invalido para a referencia '{referencia}' na linha {row}: " +
                    $"'{worksheet.Cell(row, 13).GetString().Trim()}'.");
            }

            var precoUnitario = LerDecimal(worksheet.Cell(row, 10));
            var fatorConversao = LerDecimalOpcional(worksheet.Cell(row, 9));
            var valorM2 = LerDecimalOpcional(worksheet.Cell(row, 11));
            if ((fatorConversao > 0) != (valorM2 > 0))
            {
                throw new InvalidOperationException(
                    $"Fator de conversao e valor M2 inconsistentes para '{referencia}' na linha {row}.");
            }

            var unidadeOrigem = NormalizarUnidade(worksheet.Cell(row, 5).GetString());
            // A Adama e uma representacao: compra e venda devem reproduzir
            // diretamente o PRECO 2026 do fornecedor. O VALOR M2 e apenas uma
            // referencia de conversao da tabela e nao pode substituir o preco.
            var preco = ArredondarPreco(precoUnitario);
            if (preco <= 0)
            {
                throw new InvalidOperationException(
                    $"Preco invalido para a referencia '{referencia}' na linha {row}.");
            }

            var descricaoComercial = NormalizarTexto(worksheet.Cell(row, 2).GetString());
            if (string.IsNullOrWhiteSpace(descricaoComercial))
            {
                descricaoComercial = descricaoCompleta;
            }

            var acabamento = NormalizarTexto(worksheet.Cell(row, 6).GetString());
            var linha = NormalizarTexto(worksheet.Cell(row, 7).GetString());

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = referencia,
                DescricaoCompleta = descricaoCompleta,
                DescricaoComercial = descricaoComercial,
                Grupo = linha.Equals("PRODUTO REVENDA", StringComparison.OrdinalIgnoreCase)
                    ? "PRODUTOS QUIMICOS"
                    : "REVESTIMENTOS",
                SubGrupo = acabamento,
                Marca = Marca,
                Linha = linha,
                Modelo = NormalizarValorOpcional(worksheet.Cell(row, 8).GetString()),
                Cor = NormalizarValorOpcional(worksheet.Cell(row, 4).GetString()),
                Ncm = ncm,
                UfOrigem = "SP",
                PrecoVenda = preco,
                PrecoFabrica = preco,
                Unidade = unidadeOrigem,
                QtdeEmbalagemVenda = 1,
                PesoBruto = LerDecimal(worksheet.Cell(row, 12)),
                QtdeEmbalagemCompra = 1,
                UnidFabril = unidadeOrigem,
                Observacao = string.Empty,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });

            assinaturasPorReferencia.Add(referencia, assinatura);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Adama processados.");
        Console.WriteLine(
            "[INFO] Preco de venda e preco de fabrica copiados diretamente de PRECO 2026, sem acrescimos ou conversao.");
        Console.WriteLine(
            $"[INFO] {duplicadosIdenticosIgnorados} linha(s) duplicada(s) identica(s) ignorada(s).");
        Console.WriteLine(
            "[ATENCAO] Somente NCM e UF de origem estao confirmados; os demais campos fiscais aguardam homologacao.");

        return Task.FromResult(produtos);
    }

    private static void ValidarCabecalho(IXLWorksheet worksheet)
    {
        for (var coluna = 1; coluna <= CabecalhosEsperados.Length; coluna++)
        {
            var encontrado = NormalizarComparacao(
                worksheet.Cell(LinhaCabecalho, coluna).GetString());
            var esperado = NormalizarComparacao(CabecalhosEsperados[coluna - 1]);
            if (!encontrado.Equals(esperado, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"Cabecalho Adama invalido na coluna {coluna}. " +
                    $"Esperado '{CabecalhosEsperados[coluna - 1]}', encontrado " +
                    $"'{worksheet.Cell(LinhaCabecalho, coluna).GetString().Trim()}'.");
            }
        }
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            [13] = SituacaoCampoFiscal.Confirmado,
            [14] = SituacaoCampoFiscal.Confirmado,
            [18] = SituacaoCampoFiscal.Pendente,
            [19] = SituacaoCampoFiscal.Pendente,
            [20] = SituacaoCampoFiscal.Pendente,
            [21] = SituacaoCampoFiscal.Pendente,
            [26] = SituacaoCampoFiscal.Pendente,
            [27] = SituacaoCampoFiscal.Pendente,
            [28] = SituacaoCampoFiscal.Pendente,
            [29] = SituacaoCampoFiscal.Pendente,
            [30] = SituacaoCampoFiscal.Pendente,
            [31] = SituacaoCampoFiscal.Pendente,
            [32] = SituacaoCampoFiscal.Pendente,
            [37] = SituacaoCampoFiscal.Pendente,
            [38] = SituacaoCampoFiscal.Pendente,
            [39] = SituacaoCampoFiscal.Pendente,
            [51] = SituacaoCampoFiscal.Pendente,
            [52] = SituacaoCampoFiscal.Pendente,
            [53] = SituacaoCampoFiscal.Pendente,
            [57] = SituacaoCampoFiscal.Pendente,
            [58] = SituacaoCampoFiscal.Pendente,
            [59] = SituacaoCampoFiscal.Pendente,
            [60] = SituacaoCampoFiscal.Pendente
        };
    }

    private static string CriarAssinatura(IXLWorksheet worksheet, int row)
    {
        return string.Join('|', Enumerable.Range(2, 12).Select(coluna =>
            Regex.Replace(worksheet.Cell(row, coluna).GetFormattedString().Trim(), @"\s+", " ")));
    }

    private static decimal LerDecimal(IXLCell cell)
    {
        return cell.TryGetValue<decimal>(out var valor)
            ? valor
            : DecimalParser.Parse(cell.GetFormattedString());
    }

    private static decimal LerDecimalOpcional(IXLCell cell)
    {
        var texto = cell.GetFormattedString().Trim();
        return string.IsNullOrWhiteSpace(texto) || texto == "-" ? 0 : LerDecimal(cell);
    }

    private static decimal ArredondarPreco(decimal valor)
    {
        return Math.Round(valor, 2, MidpointRounding.AwayFromZero);
    }

    private static string NormalizarUnidade(string valor)
    {
        return NormalizarComparacao(valor) switch
        {
            "PC" => "PC",
            "CX" => "CX",
            "UN" => "UN",
            var unidade => throw new InvalidOperationException(
                $"Unidade Adama nao reconhecida: '{unidade}'.")
        };
    }

    private static string SomenteDigitos(string valor)
    {
        return Regex.Replace(valor, @"\D", string.Empty);
    }

    private static string NormalizarValorOpcional(string valor)
    {
        var normalizado = NormalizarTexto(valor);
        return normalizado == "-" ? string.Empty : normalizado;
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }

    private static string NormalizarComparacao(string valor)
    {
        var normalizado = NormalizarTexto(valor)
            .Replace("²", "2", StringComparison.Ordinal)
            .Normalize(System.Text.NormalizationForm.FormD);
        normalizado = string.Concat(normalizado.Where(c =>
            System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) !=
            System.Globalization.UnicodeCategory.NonSpacingMark));
        return Regex.Replace(normalizado, @"[^A-Z0-9./()]+", string.Empty);
    }
}
