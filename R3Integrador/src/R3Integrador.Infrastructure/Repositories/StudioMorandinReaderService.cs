using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class StudioMorandinReaderService : IStudioMorandinReader
{
    private const string NomeAba = "Plan1";
    private const string Marca = "STUDIO MORANDIN";
    private const string Ncm = "69072300";
    private const string Cest = "1003000";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException($"Aba '{NomeAba}' nao encontrada na tabela Studio Morandin.");
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var linhaComercial = string.Empty;

        for (var row = 5; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 2).GetString().Trim();
            var produto = NormalizarTexto(worksheet.Cell(row, 3).GetString());

            if (string.IsNullOrWhiteSpace(referencia) || string.IsNullOrWhiteSpace(produto))
            {
                continue;
            }

            var linhaInformada = NormalizarTexto(worksheet.Cell(row, 1).GetString());
            if (!string.IsNullOrWhiteSpace(linhaInformada))
            {
                linhaComercial = linhaInformada;
            }

            var medida = NormalizarTexto(worksheet.Cell(row, 4).GetString());
            var unidade = NormalizarUnidade(worksheet.Cell(row, 6).GetString());
            var metragemCaixa = DecimalParser.Parse(worksheet.Cell(row, 7).GetFormattedString());
            var pesoCaixa = DecimalParser.Parse(worksheet.Cell(row, 9).GetFormattedString());
            var precoRevenda = DecimalParser.Parse(worksheet.Cell(row, 11).GetFormattedString());

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = referencia,
                DescricaoCompleta = CriarDescricao(produto, linhaComercial, medida),
                DescricaoComercial = produto,
                Grupo = "REVESTIMENTOS",
                SubGrupo = linhaComercial,
                Marca = Marca,
                Linha = linhaComercial,
                Modelo = medida,
                Ncm = Ncm,
                UfOrigem = string.Empty,
                // A origem chama este campo de "Preco Revenda". Ate a homologacao
                // comercial, o valor e preservado nas duas colunas sem aplicar markup.
                PrecoVenda = precoRevenda,
                PrecoFabrica = precoRevenda,
                IpiPercentual = 0,
                AliqIcmsInterna = 18,
                Iva = 81,
                Unidade = unidade,
                QtdeEmbalagemVenda = unidade == "M2" && metragemCaixa > 0 ? metragemCaixa : 1,
                PesoBruto = pesoCaixa,
                QtdeEmbalagemCompra = 1,
                UnidFabril = unidade,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = CriarObservacao(worksheet, row, Cest),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Studio Morandin processados.");
        Console.WriteLine("[ATENCAO] Saida provisoria: CST/CSOSN, CFOP, origem, ST efetiva, PIS/COFINS e IBS/CBS aguardam homologacao.");

        return Task.FromResult(produtos);
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            // NCM foi declarado expressamente no rodape da tabela de origem.
            [13] = SituacaoCampoFiscal.Confirmado,

            // Premissas atuais, ainda dependentes de homologacao fiscal.
            [18] = SituacaoCampoFiscal.Estimado,
            [20] = SituacaoCampoFiscal.Estimado,
            [21] = SituacaoCampoFiscal.Estimado,

            // Campos fiscais sem evidencia suficiente na origem atual.
            [19] = SituacaoCampoFiscal.Pendente,
            [26] = SituacaoCampoFiscal.Pendente,
            [27] = SituacaoCampoFiscal.Pendente,
            [28] = SituacaoCampoFiscal.Pendente,
            [29] = SituacaoCampoFiscal.Pendente,
            [30] = SituacaoCampoFiscal.Pendente,
            [31] = SituacaoCampoFiscal.Pendente,
            [32] = SituacaoCampoFiscal.Pendente,
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

    private static string CriarObservacao(IXLWorksheet worksheet, int row, string cest)
    {
        var espessura = worksheet.Cell(row, 5).GetFormattedString().Trim();
        var pecasM2 = worksheet.Cell(row, 8).GetFormattedString().Trim();
        var pecasCaixa = worksheet.Cell(row, 10).GetFormattedString().Trim();

        return $"Studio Morandin abril/2026 - CEST provisório {cest} - Espessura: {espessura} mm - Pecas/m2: {pecasM2} - Pecas/caixa: {pecasCaixa} - Tributacao e regra de preco pendentes de homologacao";
    }

    private static string CriarDescricao(string produto, string linha, string medida)
    {
        return string.Join(" - ", new[] { produto, linha, medida }.Where(valor => !string.IsNullOrWhiteSpace(valor)));
    }

    private static string NormalizarUnidade(string valor)
    {
        return valor.Contains("m", StringComparison.OrdinalIgnoreCase) ? "M2" : "PC";
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }
}
