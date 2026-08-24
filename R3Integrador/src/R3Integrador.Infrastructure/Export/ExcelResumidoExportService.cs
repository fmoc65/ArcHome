using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Export;

public sealed class ExcelResumidoExportService : IExcelResumidoExporter
{
    private static readonly string[] Cabecalhos =
    [
        "MARCA",
        "REFERENCIA",
        "PREÇO VENDA",
        "PREÇO DE FÁBRICA",
        "DESCRICAO",
        "UNID FABRIL",
        "MODELO",
        "COR"
    ];

    public Task ExportarAsync(IReadOnlyCollection<ProdutoErpDto> produtos, string caminhoSaida)
    {
        var diretorio = Path.GetDirectoryName(caminhoSaida);
        if (!string.IsNullOrWhiteSpace(diretorio)
            && !Directory.Exists(diretorio))
        {
            Directory.CreateDirectory(diretorio);
        }

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Plan1");

        for (var coluna = 1; coluna <= Cabecalhos.Length; coluna++)
        {
            var celula = worksheet.Cell(1, coluna);
            celula.Value = Cabecalhos[coluna - 1];
            celula.Style.Font.Bold = true;
            celula.Style.Font.FontColor = XLColor.White;
            celula.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F4E78");
        }

        var linha = 2;
        foreach (var produto in produtos)
        {
            worksheet.Cell(linha, 1).Value = produto.Marca;
            worksheet.Cell(linha, 2).Value = produto.CodigoFabrica;
            worksheet.Cell(linha, 3).Value = produto.PrecoVenda;
            worksheet.Cell(linha, 4).Value = produto.PrecoFabrica;
            worksheet.Cell(linha, 5).Value = produto.DescricaoCompleta;
            worksheet.Cell(linha, 6).Value = produto.UnidFabril;
            worksheet.Cell(linha, 7).Value = produto.Modelo;
            worksheet.Cell(linha, 8).Value = produto.Cor;
            linha++;
        }

        worksheet.Columns(3, 4).Style.NumberFormat.Format = "0.00";
        worksheet.SheetView.FreezeRows(1);
        worksheet.RangeUsed()?.SetAutoFilter();
        worksheet.Columns().AdjustToContents();

        using var fileStream = new FileStream(
            caminhoSaida,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None);
        workbook.SaveAs(fileStream);
        fileStream.Flush(true);

        Console.WriteLine($"[Sucesso] Atualizacao resumida gerada em: {caminhoSaida}");
        return Task.CompletedTask;
    }
}
