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
        if (string.IsNullOrWhiteSpace(valor)) return 0;
        valor = valor.Replace("R$", "").Replace(".", "").Replace(",", ".").Replace("-", "").Trim();
        decimal.TryParse(valor, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var resultado);
        return resultado;
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
            Marca = "ARC HOME",
            Modelo = formatoAtual.Replace(" ", string.Empty).ToUpper(),
            Faces = ParseInt(worksheet.Cell(row, 7).GetString()),
            VariacaoTonalidade = worksheet.Cell(row, 8).GetString().Trim(),
            M2Caixa = ParseDecimal(worksheet.Cell(row, 11).GetString()),
            Espessura = ParseDecimal(worksheet.Cell(row, 17).GetString())
        };
    }

    private static void PreencherPrecos(ProdutoNormalizado produto, IXLWorksheet worksheet, int row, string tabela)
    {
        produto.PrecoTabela = ParseDecimal(worksheet.Cell(row, 18).GetString());
        produto.PrecoDesconto = ParseDecimal(worksheet.Cell(row, 19).GetString());
        produto.PrecoVenda = DeveUsarPrecoDescontoComoVenda(tabela)
            ? produto.PrecoDesconto
            : ParseDecimal(worksheet.Cell(row, 20).GetString());
    }

    private static bool DeveUsarPrecoDescontoComoVenda(string tabela)
    {
        return tabela.Equals("VAREJO", StringComparison.OrdinalIgnoreCase) ||
            tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase);
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
