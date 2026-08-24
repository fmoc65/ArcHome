using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class DerossoReaderService : IDerossoReader
{
    private const string Marca = "DEROSSO";
    private const string NcmContador = "69041000";
    private const string UfOrigemContador = "SP";
    private const string CsosnContador = "500";

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

            if (string.IsNullOrWhiteSpace(sku))
            {
                continue;
            }

            var produto = NormalizarTexto(worksheet.Cell(row, 1).GetString());
            var cor = NormalizarTexto(worksheet.Cell(row, 2).GetString());
            var unidade = NormalizarUnidade(worksheet.Cell(row, 4).GetString());
            var preco = LerDecimal(worksheet.Cell(row, 5));

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
                Ncm = NcmContador,
                UfOrigem = UfOrigemContador,
                // Preserva o preco unitario da origem nas duas colunas. Para revenda,
                // a margem/markup ainda precisa ser homologada; para representacao,
                // confirmar se o preco deve ser copiado diretamente.
                PrecoVenda = preco,
                PrecoFabrica = preco,
                IpiPercentual = 0,
                AliqIcmsOrigem = 12,
                AliqIcmsInterna = 12,
                Unidade = unidade,
                QtdeEmbalagemVenda = CalcularEmbalagemVenda(unidade),
                Cst = "010",
                AliquotaCofinsCst = "01",
                AliquotaIpiCst = "49",
                AliquotaPisCst = "01",
                Csosn = CsosnContador,
                CfopDentro = "5405",
                CfopFora = "6404",
                PesoBruto = LerDecimal(worksheet.Cell(row, 11)),
                QtdeEmbalagemCompra = 1,
                PercentualSt = 0,
                UnidFabril = unidade,
                EnquadramentoIpi = "999",
                AliquotaPisOrigem = "0,65",
                AliquotaCofinsOrigem = "3",
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                AliquotaIbs = "0,1",
                AliquotaCbs = "0,9",
                ClassificacaoTributaria = "000001",
                Observacao = CriarObservacao(worksheet, row, tipoTabela),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Derosso {tipoTabela} processados.");
        Console.WriteLine("[ATENCAO] CSOSN 500 foi informado pelo usuario. Revisar com o contador por causa da retirada do NCM 6904 da ST paulista em 01/01/2026.");

        return Task.FromResult(produtos);
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            [13] = SituacaoCampoFiscal.Confirmado,
            [14] = SituacaoCampoFiscal.Confirmado,
            [30] = SituacaoCampoFiscal.Confirmado,
            [18] = SituacaoCampoFiscal.Estimado,
            [19] = SituacaoCampoFiscal.Estimado,
            [20] = SituacaoCampoFiscal.Estimado,
            [21] = SituacaoCampoFiscal.Estimado,
            [26] = SituacaoCampoFiscal.Estimado,
            [27] = SituacaoCampoFiscal.Estimado,
            [28] = SituacaoCampoFiscal.Estimado,
            [29] = SituacaoCampoFiscal.Estimado,
            [31] = SituacaoCampoFiscal.Estimado,
            [32] = SituacaoCampoFiscal.Estimado,
            [39] = SituacaoCampoFiscal.Estimado,
            [51] = SituacaoCampoFiscal.Estimado,
            [52] = SituacaoCampoFiscal.Estimado,
            [53] = SituacaoCampoFiscal.Estimado,
            [57] = SituacaoCampoFiscal.Estimado,
            [58] = SituacaoCampoFiscal.Estimado,
            [59] = SituacaoCampoFiscal.Estimado
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
        return $"Derosso {tipoTabela} 01/05/2026 - {embalagem}, {pecasEmbalagem} pecas, embalagem {dimensaoEmbalagem} - Pecas/m2 parede: {pecasParede} - Pecas/m2 piso: {pecasPiso} - NCM 69041000 e origem SP informados - CSOSN 500 informado pelo usuario - CEST 1002700/ST retirados em SP desde 01/01/2026; revisar compatibilidade do CSOSN/CFOP com o contador - {pendenciaPreco}";
    }

    private static string ObterTipoTabela(string nomeAba, string titulo)
    {
        var identificacao = $"{nomeAba} {titulo}";
        return identificacao.Contains("REVENDA", StringComparison.OrdinalIgnoreCase)
            ? "REVENDA"
            : "REPRESENTACAO";
    }

    private static decimal CalcularEmbalagemVenda(string unidade)
    {
        return unidade == "M2" ? 0 : 1;
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

    private static string NormalizarTexto(string valor) => Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
}
