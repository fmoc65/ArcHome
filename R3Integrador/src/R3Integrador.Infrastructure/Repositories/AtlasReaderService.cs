using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class AtlasReaderService : IAtlasReader
{
    private const string NomeAba = "Plan1";
    private const string Marca = "ATLAS";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException($"Aba '{NomeAba}' nao encontrada na tabela Atlas.");
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var secao = string.Empty;
        var ncm = string.Empty;
        decimal ipi = 0;
        decimal percentualSt = 0;

        for (var row = 12; row <= ultimaLinha; row++)
        {
            var coluna1 = NormalizarTexto(worksheet.Cell(row, 1).GetString());

            if (TryLerCabecalhoFiscal(coluna1, out var novaSecao, out var novoNcm))
            {
                secao = novaSecao;
                ncm = novoNcm;
                continue;
            }

            if (coluna1.Contains("DESCONTO", StringComparison.OrdinalIgnoreCase))
            {
                ipi = ExtrairPercentual(coluna1, "IPI");
                percentualSt = coluna1.Contains("ST ISENTO", StringComparison.OrdinalIgnoreCase)
                    ? 0
                    : ExtrairPercentual(coluna1, "ST");
                continue;
            }

            var produto = MapearLinhaPadrao(worksheet, row, secao, ncm, ipi, percentualSt)
                ?? MapearLinhaEspecial(worksheet, row, secao, ncm, ipi, percentualSt);

            if (produto is not null)
            {
                produtos.Add(produto);
            }
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Atlas processados.");
        Console.WriteLine("[INFO] Plan2 e Plan3 vazias foram ignoradas; regras fiscais foram herdadas por secao da Plan1.");
        Console.WriteLine("[ATENCAO] ICMS, CST/CSOSN, CFOP, PIS/COFINS e IBS/CBS aguardam contador.");

        return Task.FromResult(produtos);
    }

    private static ProdutoErpDto? MapearLinhaPadrao(
        IXLWorksheet worksheet, int row, string secao, string ncm, decimal ipi, decimal percentualSt)
    {
        var referencia = worksheet.Cell(row, 2).GetString().Trim();
        var ean = worksheet.Cell(row, 12).GetString().Trim();
        var preco = LerDecimal(worksheet.Cell(row, 5));

        if (string.IsNullOrWhiteSpace(referencia) || string.IsNullOrWhiteSpace(ean) || preco <= 0 || ncm.Length != 8)
        {
            return null;
        }

        var cor = NormalizarTexto(worksheet.Cell(row, 3).GetString());
        var formato = NormalizarTexto(worksheet.Cell(row, 4).GetString());
        var descricao = $"{secao} - {cor} - {formato}";

        return CriarProduto(
            referencia,
            ean,
            descricao,
            secao,
            formato,
            cor,
            ncm,
            preco,
            "M2",
            LerDecimal(worksheet.Cell(row, 9)),
            LerDecimal(worksheet.Cell(row, 10)),
            ipi,
            percentualSt,
            $"Codigo Atlas: {worksheet.Cell(row, 1).GetString().Trim()} - Placas/caixa: {worksheet.Cell(row, 6).GetFormattedString()} - Dimensao placa: {worksheet.Cell(row, 7).GetFormattedString()} - Area placa: {worksheet.Cell(row, 8).GetFormattedString()} - m2/pallet: {worksheet.Cell(row, 11).GetFormattedString()}");
    }

    private static ProdutoErpDto? MapearLinhaEspecial(
        IXLWorksheet worksheet, int row, string secao, string ncm, decimal ipi, decimal percentualSt)
    {
        var referencia = worksheet.Cell(row, 3).GetString().Trim();
        var preco = LerDecimal(worksheet.Cell(row, 5));

        if (string.IsNullOrWhiteSpace(referencia)
            || referencia.Equals("TIPO", StringComparison.OrdinalIgnoreCase)
            || preco <= 0
            || ncm != "69074000")
        {
            return null;
        }

        var formato = NormalizarTexto(worksheet.Cell(row, 4).GetString());
        var cores = NormalizarTexto(worksheet.Cell(row, 9).GetString());
        if (string.IsNullOrWhiteSpace(cores)) cores = NormalizarTexto(worksheet.Cell(row, 10).GetString());
        var descricao = $"{secao} - {referencia} - {formato}";

        return CriarProduto(
            referencia,
            string.Empty,
            descricao,
            secao,
            formato,
            cores,
            ncm,
            preco,
            "PC",
            1,
            LerDecimal(worksheet.Cell(row, 8)),
            ipi,
            percentualSt,
            $"Pecas/caixa: {worksheet.Cell(row, 6).GetFormattedString()} - Pecas/metro linear: {worksheet.Cell(row, 7).GetFormattedString()} - Cores: {cores}");
    }

    private static ProdutoErpDto CriarProduto(
        string referencia, string ean, string descricao, string secao, string formato,
        string cor, string ncm, decimal preco, string unidade, decimal embalagemVenda,
        decimal pesoBruto, decimal ipi, decimal percentualSt, string detalhe)
    {
        return new ProdutoErpDto
        {
            CodigoFabrica = referencia,
            CodigoBarras = ean,
            DescricaoCompleta = descricao,
            DescricaoComercial = descricao,
            Grupo = "REVESTIMENTOS",
            SubGrupo = secao,
            Marca = Marca,
            Linha = secao,
            Modelo = formato,
            Cor = cor,
            Ncm = ncm,
            UfOrigem = string.Empty,
            // Tabela de revenda 35%: preserva o preco publicado. Confirmar se ele
            // ja e custo com desconto ou se ainda exige aplicar os 35%.
            PrecoVenda = preco,
            PrecoFabrica = preco,
            DescontoPercentual = 35,
            IpiPercentual = ipi,
            Iva = percentualSt > 0 ? 81 : 0,
            Unidade = unidade,
            QtdeEmbalagemVenda = embalagemVenda > 0 ? embalagemVenda : 1,
            PesoBruto = pesoBruto,
            QtdeEmbalagemCompra = 1,
            PercentualSt = percentualSt,
            UnidFabril = unidade,
            EstoqueMinimo = 0,
            EstoqueMaximo = 0,
            Observacao = $"Atlas revenda 35% maio/2026 - {detalhe} - NCM/IPI/ST informados pela secao da origem - regra de preco e demais tributos pendentes",
            SituacaoCamposFiscais = CriarSituacaoCamposFiscais(percentualSt)
        };
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais(decimal percentualSt)
    {
        var situacoes = new Dictionary<int, SituacaoCampoFiscal>
        {
            [13] = SituacaoCampoFiscal.Confirmado,
            [18] = SituacaoCampoFiscal.Confirmado,
            [39] = SituacaoCampoFiscal.Confirmado,
            [19] = SituacaoCampoFiscal.Pendente,
            [20] = SituacaoCampoFiscal.Pendente,
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

        situacoes[21] = percentualSt > 0
            ? SituacaoCampoFiscal.Estimado
            : SituacaoCampoFiscal.Pendente;

        return situacoes;
    }

    private static bool TryLerCabecalhoFiscal(string texto, out string secao, out string ncm)
    {
        var match = Regex.Match(texto, @"NCM\s*-?\s*(\d{4})\.(\d{2})\.(\d{2})", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            secao = string.Empty;
            ncm = string.Empty;
            return false;
        }

        ncm = string.Concat(match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value);
        secao = Regex.Replace(texto[..match.Index], @"\s*-\s*$", string.Empty).Trim();
        return true;
    }

    private static decimal ExtrairPercentual(string texto, string campo)
    {
        var match = Regex.Match(texto, $@"{campo}\s*(\d+[,.]?\d*)%", RegexOptions.IgnoreCase);
        return match.Success ? DecimalParser.Parse(match.Groups[1].Value) : 0;
    }

    private static decimal LerDecimal(IXLCell cell)
    {
        return cell.TryGetValue<decimal>(out var valor) ? valor : DecimalParser.Parse(cell.GetFormattedString());
    }

    private static string NormalizarTexto(string valor) => Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
}
