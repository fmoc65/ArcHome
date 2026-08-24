using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Repositories;

public class ExcelReaderService : IExcelReader
{
    public async Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo, string tabela)
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
        string linhaModeloAtual = string.Empty;

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int row = 1; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 2).GetString().Trim();

            if (DeveIgnorarLinha(referencia, tabela))
                continue;

            var formato = worksheet.Cell(row, 1).GetString().Trim();
            var linhaModelo = worksheet.Cell(row, 3).GetString().Trim();
            var superficie = NormalizarSuperficie(worksheet.Cell(row, 6).GetString(), tabela);

            formatoAtual = ManterValorMesclado(formato, formatoAtual);
            linhaModeloAtual = ManterValorMesclado(linhaModelo, linhaModeloAtual);

            var produto = CriarProduto(worksheet, row, tabela, referencia, formatoAtual, linhaModeloAtual, superficie);
            PreencherPrecos(produto, worksheet, row, tabela);

            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos processados com sucesso na aba {tabela}.");

        return await Task.FromResult(produtos);
    }

    private static decimal ParseDecimal(string valor)
    {
        return DecimalParser.Parse(valor);
    }

    private static int ParseInt(string valor)
    {
        int.TryParse(valor, out var resultado);
        return resultado;
    }

    private static bool DeveIgnorarLinha(string referencia, string tabela)
    {
        if (string.IsNullOrWhiteSpace(referencia))
            return true;

        if (EhCabecalho(referencia))
            return true;

        return !tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase) &&
            !referencia.All(char.IsDigit);
    }

    private static bool EhCabecalho(string referencia)
    {
        return referencia.Equals("Ref.", StringComparison.OrdinalIgnoreCase) ||
            referencia.Equals("VINÍLICOS", StringComparison.OrdinalIgnoreCase) ||
            referencia.Equals("VINILICOS", StringComparison.OrdinalIgnoreCase) ||
            referencia.Equals("LASTRAS", StringComparison.OrdinalIgnoreCase) ||
            referencia.Equals("Formato", StringComparison.OrdinalIgnoreCase);
    }

    private static string ManterValorMesclado(string valorAtual, string valorAnterior)
    {
        return string.IsNullOrWhiteSpace(valorAtual) ? valorAnterior : valorAtual;
    }

    private static ProdutoNormalizado CriarProduto(
        IXLWorksheet worksheet,
        int row,
        string tabela,
        string referencia,
        string formatoAtual,
        string linhaModeloAtual,
        string superficie)
    {
        return new ProdutoNormalizado
        {
            TipoTabela = tabela,
            Formato = formatoAtual,
            Referencia = referencia,
            Linha = linhaModeloAtual,
            Colecao = worksheet.Cell(row, 4).GetString().Trim(),
            Cor = worksheet.Cell(row, 5).GetString().Trim(),
            Superficie = superficie,
            Grupo = "PORCELANATO",
            SubGrupo = superficie.ToUpper(),
            Marca = "VILLAGRES",
            Modelo = formatoAtual.Replace(" ", string.Empty).ToUpper(),
            Faces = ParseInt(worksheet.Cell(row, 7).GetString()),
            VariacaoTonalidade = worksheet.Cell(row, 8).GetString().Trim(),
            M2Caixa = ParseDecimal(worksheet.Cell(row, 11).GetString()),
            Espessura = ParseDecimal(worksheet.Cell(row, 17).GetString()),
            PesoBrutoM2 = ParseDecimal(worksheet.Cell(row, 15).GetString())
        };
    }

    private static void PreencherPrecos(ProdutoNormalizado produto, IXLWorksheet worksheet, int row, string tabela)
    {
        produto.PrecoDesconto = ParseDecimal(worksheet.Cell(row, 19).GetString());
        if (tabela.Equals("VAREJO", StringComparison.OrdinalIgnoreCase))
        {
            produto.PrecoTabela = CalcularPrecoFinalVarejo(
                produto.PrecoDesconto,
                produto.Referencia);
            produto.PrecoVenda = produto.PrecoTabela;
            return;
        }

        produto.PrecoTabela = tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase)
            ? CalcularCustoFinalRevenda(produto.PrecoDesconto)
            : ParseDecimal(worksheet.Cell(row, 18).GetString());
        produto.PrecoVenda = tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase)
            ? CalcularPrecoVendaRevenda(produto.PrecoTabela)
            : ParseDecimal(worksheet.Cell(row, 20).GetString());
    }

    private static decimal CalcularPrecoFinalVarejo(decimal precoDesconto, string referencia)
    {
        const decimal percentualSt = 9.86m;
        var percentualIpi = referencia.Equals("120003", StringComparison.OrdinalIgnoreCase)
            ? 0m
            : 0.65m;
        var fator = 1m + percentualIpi / 100m + percentualSt / 100m;
        return Math.Round(precoDesconto * fator, 2, MidpointRounding.AwayFromZero);
    }

    private static decimal CalcularCustoFinalRevenda(decimal precoDesconto)
    {
        return Math.Round(precoDesconto * 1.1051m + 1.50m, 2);
    }

    private static decimal CalcularPrecoVendaRevenda(decimal custoFinal)
    {
        return Math.Round(custoFinal * 1.75m, 2);
    }

    private static string NormalizarSuperficie(string valor, string tabela)
    {
        var superficie = valor.Trim();

        if (tabela.Equals("VAREJO", StringComparison.OrdinalIgnoreCase) &&
            superficie.Equals("NATURAL SENSE UP", StringComparison.OrdinalIgnoreCase))
        {
            return "NATURAL SENSEUP";
        }

        return superficie;
    }
}
