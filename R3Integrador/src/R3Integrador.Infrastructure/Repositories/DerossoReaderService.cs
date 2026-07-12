using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class DerossoReaderService : IDerossoReader
{
    private const string Marca = "DEROSSO";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);
        var worksheet = workbook.Worksheet(1);
        var tipoTabela = ObterTipoTabela(worksheet.Name, worksheet.Cell(1, 1).GetString());
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 4; row <= ultimaLinha; row++)
        {
            var sku = worksheet.Cell(row, 3).GetString().Trim();
            var ncm = SomenteDigitos(worksheet.Cell(row, 7).GetString());

            if (string.IsNullOrWhiteSpace(sku) || ncm.Length != 8)
            {
                continue;
            }

            var produto = NormalizarTexto(worksheet.Cell(row, 1).GetString());
            var cor = NormalizarTexto(worksheet.Cell(row, 2).GetString());
            var unidade = NormalizarUnidade(worksheet.Cell(row, 4).GetString());
            var preco = LerDecimal(worksheet.Cell(row, 5));
            var embalagensPorM2 = LerDecimal(worksheet.Cell(row, 13));

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = sku,
                DescricaoCompleta = $"{produto} - {cor}",
                DescricaoComercial = produto,
                Grupo = "REVESTIMENTOS",
                SubGrupo = produto,
                Marca = Marca,
                Linha = produto,
                Modelo = NormalizarTexto(worksheet.Cell(row, 8).GetString()),
                Cor = cor,
                Ncm = ncm,
                UfOrigem = string.Empty,
                // Preserva o preco unitario da origem nas duas colunas. Para revenda,
                // a margem/markup ainda precisa ser homologada; para representacao,
                // confirmar se o preco deve ser copiado diretamente.
                PrecoVenda = preco,
                PrecoFabrica = preco,
                AliqIcmsInterna = 12,
                Unidade = unidade,
                QtdeEmbalagemVenda = CalcularEmbalagemVenda(unidade, embalagensPorM2),
                PesoBruto = LerDecimal(worksheet.Cell(row, 11)),
                QtdeEmbalagemCompra = 1,
                PercentualSt = 0,
                UnidFabril = unidade,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = CriarObservacao(worksheet, row, tipoTabela),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Derosso {tipoTabela} processados.");
        Console.WriteLine("[ATENCAO] ICMS 12% e ausencia de ST sao premissas; demais campos fiscais aguardam contador.");

        return Task.FromResult(produtos);
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            [13] = SituacaoCampoFiscal.Confirmado,
            [20] = SituacaoCampoFiscal.Estimado,
            [39] = SituacaoCampoFiscal.Estimado,
            [18] = SituacaoCampoFiscal.Pendente,
            [19] = SituacaoCampoFiscal.Pendente,
            [21] = SituacaoCampoFiscal.Pendente,
            [26] = SituacaoCampoFiscal.Pendente,
            [27] = SituacaoCampoFiscal.Pendente,
            [28] = SituacaoCampoFiscal.Pendente,
            [29] = SituacaoCampoFiscal.Pendente,
            [30] = SituacaoCampoFiscal.Pendente,
            [31] = SituacaoCampoFiscal.Pendente,
            [32] = SituacaoCampoFiscal.Pendente,
            [51] = SituacaoCampoFiscal.Pendente,
            [52] = SituacaoCampoFiscal.Pendente,
            [53] = SituacaoCampoFiscal.Pendente,
            [57] = SituacaoCampoFiscal.Pendente,
            [58] = SituacaoCampoFiscal.Pendente,
            [59] = SituacaoCampoFiscal.Pendente,
            [60] = SituacaoCampoFiscal.Pendente
        };
    }

    private static string CriarObservacao(IXLWorksheet worksheet, int row, string tipoTabela)
    {
        var embalagem = NormalizarTexto(worksheet.Cell(row, 9).GetString());
        var pecasEmbalagem = worksheet.Cell(row, 10).GetFormattedString().Trim();
        var dimensaoEmbalagem = NormalizarTexto(worksheet.Cell(row, 12).GetString());
        var pecasParede = worksheet.Cell(row, 14).GetFormattedString().Trim();
        var pecasPiso = worksheet.Cell(row, 15).GetFormattedString().Trim();

        var pendenciaPreco = tipoTabela == "REVENDA"
            ? "regra de margem/markup da revenda pendente"
            : "regra comercial da representacao pendente";
        return $"Derosso {tipoTabela} 01/05/2026 - {embalagem}, {pecasEmbalagem} pecas, embalagem {dimensaoEmbalagem} - Pecas/m2 parede: {pecasParede} - Pecas/m2 piso: {pecasPiso} - CEST 1002700/ST revogados em SP desde 01/01/2026 - {pendenciaPreco} - confirmar NCM, acabamento, ICMS e tributacao com contador";
    }

    private static string ObterTipoTabela(string nomeAba, string titulo)
    {
        var identificacao = $"{nomeAba} {titulo}";
        return identificacao.Contains("REVENDA", StringComparison.OrdinalIgnoreCase)
            ? "REVENDA"
            : "REPRESENTACAO";
    }

    private static decimal CalcularEmbalagemVenda(string unidade, decimal embalagensPorM2)
    {
        return unidade == "M2" && embalagensPorM2 > 0
            ? Math.Round(1 / embalagensPorM2, 4)
            : 1;
    }

    private static string NormalizarUnidade(string valor)
    {
        var unidade = NormalizarTexto(valor);
        if (unidade.Contains("M²") || unidade == "M2") return "M2";
        if (unidade is "PÇ" or "PC") return "PC";
        if (unidade == "KG") return "KG";
        return unidade;
    }

    private static decimal LerDecimal(IXLCell cell)
    {
        return cell.TryGetValue<decimal>(out var valor)
            ? valor
            : DecimalParser.Parse(cell.GetFormattedString());
    }

    private static string SomenteDigitos(string valor) => Regex.Replace(valor, @"\D", string.Empty);
    private static string NormalizarTexto(string valor) => Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
}
