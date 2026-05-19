using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Export;

public class ExcelExportService : IExcelExporter
{
    public async Task ExportarAsync(
        List<ProdutoErpDto> produtos,
        string caminhoSaida)
    {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("IMPORTACAO_ERP");

        CriarCabecalho(worksheet);

        int linha = 2;

        foreach (var produto in produtos)
        {
            worksheet.Cell(linha, 1).Value = produto.CodigoInterno;
            worksheet.Cell(linha, 2).Value = produto.CodigoFabrica;
            worksheet.Cell(linha, 3).Value = produto.DescricaoCompleta;
            worksheet.Cell(linha, 4).Value = produto.Grupo;       // Agora numérico (Ex: 10, 20) conforme o EnumNormalizer
            worksheet.Cell(linha, 5).Value = produto.SubGrupo;    // Agora numérico (Ex: 11, 21) conforme o EnumNormalizer
            worksheet.Cell(linha, 6).Value = produto.Marca;
            worksheet.Cell(linha, 7).Value = produto.PrecoVenda;
            worksheet.Cell(linha, 8).Value = produto.Unidade;
            worksheet.Cell(linha, 9).Value = produto.Ncm;
            worksheet.Cell(linha, 10).Value = produto.Cst;
            worksheet.Cell(linha, 11).Value = produto.Cfop;
            worksheet.Cell(linha, 12).Value = produto.PesoBruto;
            worksheet.Cell(linha, 13).Value = produto.Espessura;
            worksheet.Cell(linha, 14).Value = produto.PrecoFracionado;

            linha++;
        }

        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(caminhoSaida);

        Console.WriteLine();
        Console.WriteLine("Arquivo exportado com sucesso para o ERP:");
        Console.WriteLine(caminhoSaida);

        await Task.CompletedTask;
    }

    private static void CriarCabecalho(IXLWorksheet worksheet)
    {
        // Alinhado com as colunas exigidas pelo layout_importacao_produto.xlsx do seu ERP
        worksheet.Cell(1, 1).Value = "CÓDIGO INTERNO";
        worksheet.Cell(1, 2).Value = "CÓDIGO FÁBRICA";
        worksheet.Cell(1, 3).Value = "DESCRIÇÃO COMPLETA";
        worksheet.Cell(1, 4).Value = "GRUPO";
        worksheet.Cell(1, 5).Value = "SUBGRUPO";
        worksheet.Cell(1, 6).Value = "MARCA";
        worksheet.Cell(1, 7).Value = "PREÇO VENDA";
        worksheet.Cell(1, 8).Value = "UNIDADE";
        worksheet.Cell(1, 9).Value = "NCM";
        worksheet.Cell(1, 10).Value = "CST";
        worksheet.Cell(1, 11).Value = "CFOP DENTRO";
        worksheet.Cell(1, 12).Value = "PESOBRUTO";
        worksheet.Cell(1, 13).Value = "ESPESSURA";
        worksheet.Cell(1, 14).Value = "PRECO_FRACIONADO";
    }
}