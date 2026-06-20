using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public class RocaReaderService : IRocaReader
{
    private const string GrupoPadrao = "LOUCAS E METAIS";
    private const decimal AliquotaIcmsInternaPadrao = 18m;

    public async Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);
        var worksheet = workbook.Worksheets.First();
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 2; row <= ultimaLinha; row++)
        {
            var codigoFabrica = worksheet.Cell(row, 4).GetString().Trim();
            var descricaoComercial = worksheet.Cell(row, 6).GetString().Trim();

            if (!EhProdutoValido(codigoFabrica, descricaoComercial))
            {
                continue;
            }

            var precoTabela = ParseDecimal(worksheet.Cell(row, 10).GetFormattedString());
            var precoCampanha = ParseDecimal(worksheet.Cell(row, 12).GetFormattedString());
            var precoFinal = precoCampanha > 0 ? precoCampanha : precoTabela;
            var linha = NormalizarTexto(worksheet.Cell(row, 18).GetString());
            var segmento = NormalizarTexto(worksheet.Cell(row, 19).GetString());
            var cest = SomenteDigitos(worksheet.Cell(row, 36).GetString());

            produtos.Add(new ProdutoErpDto
            {
                CodigoInterno = string.Empty,
                CodigoFabrica = codigoFabrica,
                CodigoBarras = worksheet.Cell(row, 5).GetString().Trim(),
                DescricaoCompleta = ObterDescricaoCompleta(worksheet, row, descricaoComercial),
                DescricaoComercial = NormalizarTexto(descricaoComercial),
                Grupo = GrupoPadrao,
                SubGrupo = string.IsNullOrWhiteSpace(linha) ? "SEM LINHA" : linha,
                Marca = NormalizarMarca(worksheet.Cell(row, 17).GetString()),
                Linha = linha,
                Modelo = segmento,
                Voltagem = string.Empty,
                Cor = NormalizarTexto(worksheet.Cell(row, 20).GetString()),
                Ncm = SomenteDigitos(worksheet.Cell(row, 29).GetString()),
                UfOrigem = "SP",
                PrecoVenda = precoFinal,
                PrecoFabrica = precoFinal,
                DescontoPercentual = CalcularDescontoPercentual(precoTabela, precoFinal),
                IpiPercentual = ParseDecimal(worksheet.Cell(row, 14).GetFormattedString()),
                AliqIcmsOrigem = ParseDecimal(worksheet.Cell(row, 15).GetFormattedString()),
                AliqIcmsInterna = AliquotaIcmsInternaPadrao,
                Iva = ParseDecimal(worksheet.Cell(row, 16).GetFormattedString()),
                FreteReais = 0,
                FretePercentual = 0,
                Unidade = "UN",
                QtdeEmbalagemVenda = ObterQuantidadeEmbalagem(worksheet, row),
                Cst = "060",
                AliquotaCofinsCst = "01",
                AliquotaIpiCst = "49",
                AliquotaPisCst = "01",
                Csosn = "500",
                CfopDentro = "5405",
                CfopFora = "6404",
                PesoLiquido = ParseDecimal(worksheet.Cell(row, 24).GetFormattedString()),
                PesoBruto = ParseDecimal(worksheet.Cell(row, 25).GetFormattedString()),
                QtdeEmbalagemCompra = 1,
                ValorPi = 0,
                AliquotaCofins = 0,
                AliquotaPis = 0,
                PercentualSt = ParseDecimal(worksheet.Cell(row, 16).GetFormattedString()),
                UnidFabril = "UN",
                Observacao = CriarObservacao(worksheet, row, cest),
                DiferencaIcms = 0,
                ReducaoBaseIcms = 0,
                ReducaoBaseSt = 0,
                    EnquadramentoIpi = "0",
                    AliquotaIbs = "0",
                    AliquotaCbs = "0",
                    ClassificacaoTributaria = "0"
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos ROCA/CELITE processados com sucesso.");

        return await Task.FromResult(produtos);
    }

    private static bool EhProdutoValido(string codigoFabrica, string descricao)
    {
        return !string.IsNullOrWhiteSpace(codigoFabrica)
            && !string.IsNullOrWhiteSpace(descricao)
            && !codigoFabrica.Equals("Código", StringComparison.OrdinalIgnoreCase);
    }

    private static string ObterDescricaoCompleta(IXLWorksheet worksheet, int row, string descricaoComercial)
    {
        var descricaoCompleta = worksheet.Cell(row, 31).GetString().Trim();
        return NormalizarTexto(string.IsNullOrWhiteSpace(descricaoCompleta) ? descricaoComercial : descricaoCompleta);
    }

    private static string NormalizarMarca(string valor)
    {
        return valor.Trim().ToUpperInvariant() switch
        {
            "CE" => "CELITE",
            "RO" => "ROCA",
            _ => NormalizarTexto(valor)
        };
    }

    private static decimal ObterQuantidadeEmbalagem(IXLWorksheet worksheet, int row)
    {
        var quantidade = ParseDecimal(worksheet.Cell(row, 28).GetFormattedString());
        return quantidade > 0 ? quantidade : 1;
    }

    private static string CriarObservacao(IXLWorksheet worksheet, int row, string cest)
    {
        var origem = worksheet.Cell(row, 30).GetString().Trim();
        var dimensoes = $"Dimensoes produto LxPxA: {worksheet.Cell(row, 21).GetFormattedString()}x{worksheet.Cell(row, 22).GetFormattedString()}x{worksheet.Cell(row, 23).GetFormattedString()}";
        var embalagem = $"Embalagem LxPxA: {worksheet.Cell(row, 34).GetFormattedString()}x{worksheet.Cell(row, 35).GetFormattedString()}x{worksheet.Cell(row, 33).GetFormattedString()}";
        var observacao = $"Importado via R3Integrador - Tabela ROCA - {dimensoes} - {embalagem}";

        if (!string.IsNullOrWhiteSpace(cest))
        {
            observacao += $" - CEST: {cest}";
        }

        if (!string.IsNullOrWhiteSpace(origem))
        {
            observacao += $" - Origem: {origem}";
        }

        return observacao;
    }

    private static decimal CalcularDescontoPercentual(decimal precoTabela, decimal precoFinal)
    {
        if (precoTabela <= 0 || precoFinal <= 0 || precoFinal >= precoTabela)
        {
            return 0;
        }

        return Math.Round((precoTabela - precoFinal) / precoTabela * 100, 2);
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }

    private static string SomenteDigitos(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\D", string.Empty);
    }

    private static decimal ParseDecimal(string valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return 0;
        }

        valor = valor.Replace("R$", "")
            .Replace("%", "")
            .Replace(".", "")
            .Replace(",", ".")
            .Replace("-", "")
            .Trim();

        decimal.TryParse(
            valor,
            NumberStyles.Any,
            CultureInfo.InvariantCulture,
            out var resultado);

        return resultado;
    }
}



