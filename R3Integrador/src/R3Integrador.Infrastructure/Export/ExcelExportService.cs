using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Export;

public class ExcelExportService : IExcelExporter
{
    public async Task ExportarAsync(
        List<ProdutoNormalizado> produtos,
        string caminhoSaida)
    {
        using var workbook = new XLWorkbook();

        var worksheet =
            workbook.Worksheets.Add("IMPORTACAO_ERP");

        CriarCabecalho(worksheet);

        int linha = 2;

        foreach (var produto in produtos)
        {
            worksheet.Cell(linha, 1).Value =
                produto.Referencia;

            worksheet.Cell(linha, 2).Value =
                $"{produto.Linha} {produto.Formato}";

            worksheet.Cell(linha, 3).Value =
                produto.Colecao;

            worksheet.Cell(linha, 4).Value =
                produto.PrecoTabela;

            worksheet.Cell(linha, 5).Value =
                produto.PrecoDesconto;

            worksheet.Cell(linha, 6).Value =
                produto.M2Caixa;

            worksheet.Cell(linha, 7).Value =
                produto.Espessura;

            linha++;
        }

        worksheet.Columns().AdjustToContents();

        workbook.SaveAs(caminhoSaida);

        Console.WriteLine();
        Console.WriteLine($"Arquivo exportado:");
        Console.WriteLine(caminhoSaida);

        await Task.CompletedTask;
    }

    private static void CriarCabecalho(
        IXLWorksheet worksheet)
    {
        worksheet.Cell(1, 1).Value = "CODIGO";
        worksheet.Cell(1, 2).Value = "DESCRICAO";
        worksheet.Cell(1, 3).Value = "COLECAO";
        worksheet.Cell(1, 4).Value = "PRECO_TABELA";
        worksheet.Cell(1, 5).Value = "PRECO_DESCONTO";
        worksheet.Cell(1, 6).Value = "M2_CAIXA";
        worksheet.Cell(1, 7).Value = "ESPESSURA";

        var range =
            worksheet.Range(1, 1, 1, 7);

        range.Style.Font.Bold = true;
    }
}