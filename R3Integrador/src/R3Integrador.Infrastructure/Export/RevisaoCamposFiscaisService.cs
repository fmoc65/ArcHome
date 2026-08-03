using ClosedXML.Excel;

namespace R3Integrador.Infrastructure.Export;

public sealed class RevisaoCamposFiscaisService
{
    private static readonly IReadOnlyDictionary<int, string> CamposObrigatorios = new Dictionary<int, string>
    {
        [19] = "ALIQICMSORIGEM",
        [20] = "ALIQICMSINTERNA",
        [26] = "CST",
        [27] = "ALIQUOTA COFINS CST",
        [28] = "ALIQUOTA IPI CST",
        [29] = "ALIQUOTA PIS CST",
        [30] = "CSOSN",
        [51] = "ENQUADRAMENTO IPI",
        [57] = "ALIQUOTA IBS",
        [58] = "ALIQUOTA CBS",
        [59] = "CLASSIFICACAO TRIBUTARIA"
    };

    public RevisaoCamposFiscaisResultado Revisar(string caminhoArquivo)
    {
        using var workbook = new XLWorkbook(caminhoArquivo);
        if (!workbook.TryGetWorksheet("IMPORTACAO_ERP", out var worksheet))
        {
            throw new InvalidOperationException("Aba 'IMPORTACAO_ERP' nao encontrada na planilha.");
        }

        var faltasPorCampo = CamposObrigatorios.Values.ToDictionary(campo => campo, _ => 0);
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 1;

        for (var linha = 2; linha <= ultimaLinha; linha++)
        {
            // Uma linha sem codigo de fabrica nao representa produto para importacao.
            if (string.IsNullOrWhiteSpace(worksheet.Cell(linha, 2).GetFormattedString()))
            {
                continue;
            }

            foreach (var (coluna, campo) in CamposObrigatorios)
            {
                var celula = worksheet.Cell(linha, coluna);
                if (!string.IsNullOrWhiteSpace(celula.GetFormattedString()))
                {
                    continue;
                }

                AplicarPendente(celula);
                AplicarCabecalhoPendente(worksheet.Cell(1, coluna));
                faltasPorCampo[campo]++;
            }
        }

        var diretorio = Path.GetDirectoryName(caminhoArquivo) ?? Directory.GetCurrentDirectory();
        var nomeSemExtensao = Path.GetFileNameWithoutExtension(caminhoArquivo);
        var caminhoSaida = Path.Combine(diretorio, $"{nomeSemExtensao}_PENDENCIAS_FISCAIS.xlsx");
        workbook.SaveAs(caminhoSaida);

        return new RevisaoCamposFiscaisResultado(caminhoSaida, faltasPorCampo);
    }

    private static void AplicarPendente(IXLCell celula)
    {
        celula.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFC7CE");
        celula.Style.Font.FontColor = XLColor.FromHtml("#9C0006");
    }

    private static void AplicarCabecalhoPendente(IXLCell cabecalho)
    {
        cabecalho.Style.Fill.BackgroundColor = XLColor.FromHtml("#C00000");
        cabecalho.Style.Font.FontColor = XLColor.White;
    }
}

public sealed record RevisaoCamposFiscaisResultado(
    string CaminhoSaida,
    IReadOnlyDictionary<string, int> FaltasPorCampo);
