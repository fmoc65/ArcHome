using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Repositories;

public class ExcelReaderService : IExcelReader
{
    public async Task<List<ProdutoNormalizado>> LerAsync(
        string caminhoArquivo,
        string tabela)
    {
        var produtos = new List<ProdutoNormalizado>();

        using var workbook = new XLWorkbook(caminhoArquivo);

        var worksheet = workbook.Worksheet(tabela);

        if (worksheet == null)
        {
            Console.WriteLine($"Aba {tabela} não encontrada.");

            return produtos;
        }

        string formatoAtual = string.Empty;

        var ultimaLinha =
            worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int row = 1; row <= ultimaLinha; row++)
        {
            var referencia =
                worksheet.Cell(row, 2)
                    .GetString()
                    .Trim();

            // ignora linhas vazias
            if (string.IsNullOrWhiteSpace(referencia))
                continue;

            // ignora cabeçalhos/textos
            if (!referencia.All(char.IsDigit))
                continue;

            var formato =
                worksheet.Cell(row, 1)
                    .GetString()
                    .Trim();

            // mantém valor do merge
            if (!string.IsNullOrWhiteSpace(formato))
            {
                formatoAtual = formato;
            }

            var produto = new ProdutoNormalizado
            {
                TipoTabela = worksheet.Name,

                Formato = formatoAtual,

                Referencia = referencia,

                Linha =
                    worksheet.Cell(row, 3)
                        .GetString()
                        .Trim(),

                Colecao =
                    worksheet.Cell(row, 4)
                        .GetString()
                        .Trim(),

                Cor =
                    worksheet.Cell(row, 5)
                        .GetString()
                        .Trim(),

                Superficie =
                    worksheet.Cell(row, 6)
                        .GetString()
                        .Trim(),

                Faces = ParseInt(
                    worksheet.Cell(row, 7)
                        .GetString()),

                VariacaoTonalidade =
                    worksheet.Cell(row, 8)
                        .GetString()
                        .Trim(),

                M2Caixa = ParseDecimal(
                    worksheet.Cell(row, 11)
                        .GetString()),

                Espessura = ParseDecimal(
                    worksheet.Cell(row, 17)
                        .GetString()),

                PrecoTabela = ParseDecimal(
                    worksheet.Cell(row, 18)
                        .GetString()),

                PrecoDesconto = ParseDecimal(
                    worksheet.Cell(row, 19)
                        .GetString())
            };

            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine(
            $"{produtos.Count} produtos processados.");

        return await Task.FromResult(produtos);
    }

    private static decimal ParseDecimal(string valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
            return 0;

        valor = valor
            .Replace("R$", "")
            .Replace(".", "")
            .Replace(",", ".")
            .Trim();

        decimal.TryParse(
            valor,
            System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture,
            out var resultado);

        return resultado;
    }

    private static int ParseInt(string valor)
    {
        int.TryParse(valor, out var resultado);

        return resultado;
    }
}