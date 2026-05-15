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
                produto.TipoTabela;

            worksheet.Cell(linha, 2).Value =
                produto.Formato;

            worksheet.Cell(linha, 3).Value =
                produto.Referencia;

            worksheet.Cell(linha, 4).Value =
                produto.Linha;

            worksheet.Cell(linha, 5).Value =
                produto.Colecao;

            worksheet.Cell(linha, 6).Value =
                produto.Cor;

            worksheet.Cell(linha, 7).Value =
                produto.Superficie;

            worksheet.Cell(linha, 8).Value =
                produto.Faces;

            worksheet.Cell(linha, 9).Value =
                produto.VariacaoTonalidade;

            worksheet.Cell(linha, 10).Value =
                produto.PrecoTabela;

            worksheet.Cell(linha, 11).Value =
                produto.PrecoDesconto;

            worksheet.Cell(linha, 12).Value =
                produto.Espessura;

            worksheet.Cell(linha, 13).Value =
                produto.M2Caixa;

            linha++;
        }

        worksheet.Columns().AdjustToContents();

        workbook.SaveAs(caminhoSaida);

        Console.WriteLine();
        Console.WriteLine("Arquivo exportado:");
        Console.WriteLine(caminhoSaida);

        await Task.CompletedTask;
    }

    private static void CriarCabecalho(
        IXLWorksheet worksheet)
    {
        worksheet.Cell(1, 1).Value = "TIPO_TABELA";
        worksheet.Cell(1, 2).Value = "FORMATO";
        worksheet.Cell(1, 3).Value = "REFERENCIA";
        worksheet.Cell(1, 4).Value = "LINHA";
        worksheet.Cell(1, 5).Value = "COLECAO";
        worksheet.Cell(1, 6).Value = "COR";
        worksheet.Cell(1, 7).Value = "SUPERFICIE";
        worksheet.Cell(1, 8).Value = "FACES";
        worksheet.Cell(1, 9).Value = "VARIACAO_TONALIDADE";
        worksheet.Cell(1, 10).Value = "PRECO_TABELA";
        worksheet.Cell(1, 11).Value = "PRECO_DESCONTO";
        worksheet.Cell(1, 12).Value = "ESPESSURA";
        worksheet.Cell(1, 13).Value = "M2_CAIXA";

        var range =
            worksheet.Range(1, 1, 1, 13);

        range.Style.Font.Bold = true;
    }

}