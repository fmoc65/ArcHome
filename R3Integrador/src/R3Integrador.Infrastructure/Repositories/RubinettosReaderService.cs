using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public class RubinettosReaderService : IRubinettosReader
{
    public async Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        foreach (var worksheet in workbook.Worksheets)
        {
            var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

            for (var row = 2; row <= ultimaLinha; row++)
            {
                var codigoFabrica = worksheet.Cell(row, 7).GetString().Trim();
                var descricao = worksheet.Cell(row, 8).GetString().Trim();
                var marca = NormalizarMarca(worksheet.Cell(row, 14).GetString());

                if (!EhProdutoValido(codigoFabrica, descricao, marca))
                {
                    continue;
                }

                var linha = NormalizarTexto(worksheet.Cell(row, 13).GetString());
                var unidade = NormalizarUnidade(worksheet.Cell(row, 22).GetString());
                var ncm = SomenteDigitos(worksheet.Cell(row, 29).GetString());
                var cest = worksheet.Cell(row, 30).GetString().Trim();

                produtos.Add(new ProdutoErpDto
                {
                    CodigoInterno = string.Empty,
                    CodigoFabrica = codigoFabrica,
                    CodigoBarras = worksheet.Cell(row, 31).GetString().Trim(),
                    DescricaoCompleta = NormalizarTexto(descricao),
                    DescricaoComercial = NormalizarTexto(descricao),
                    Grupo = "METAIS",
                    SubGrupo = string.IsNullOrWhiteSpace(linha) ? "METAIS" : linha,
                    Marca = marca,
                    Linha = linha,
                    Modelo = NormalizarTexto(worksheet.Cell(row, 5).GetString()),
                    Voltagem = "N/A",
                    Cor = NormalizarTexto(worksheet.Cell(row, 10).GetString()),
                    Ncm = ncm,
                    UfOrigem = "SP",
                    PrecoVenda = ParseDecimal(worksheet.Cell(row, 15).GetFormattedString()),
                    PrecoFabrica = ParseDecimal(worksheet.Cell(row, 16).GetFormattedString()),
                    DescontoPercentual = 0,
                    IpiPercentual = ParseDecimal(worksheet.Cell(row, 24).GetFormattedString()),
                    AliqIcmsOrigem = ParseDecimal(worksheet.Cell(row, 25).GetFormattedString()),
                    AliqIcmsInterna = ParseDecimal(worksheet.Cell(row, 27).GetFormattedString()),
                    Iva = ParseDecimal(worksheet.Cell(row, 26).GetFormattedString()),
                    FreteReais = 0,
                    FretePercentual = 0,
                    Unidade = unidade,
                    QtdeEmbalagemVenda = 1,
                    Cst = "060",
                    AliquotaCofinsCst = "01",
                    AliquotaIpiCst = "49",
                    AliquotaPisCst = "01",
                    Csosn = "500",
                    CfopDentro = "5405",
                    CfopFora = "6404",
                    PesoLiquido = 0,
                    PesoBruto = 0,
                    QtdeEmbalagemCompra = 1,
                    ValorPi = 0,
                    AliquotaCofins = 0,
                    AliquotaPis = 0,
                    PercentualSt = ParseDecimal(worksheet.Cell(row, 26).GetFormattedString()),
                    UnidFabril = unidade,
                    Observacao = CriarObservacao(worksheet.Name, cest),
                    DiferencaIcms = 0,
                    ReducaoBaseIcms = 0,
                    ReducaoBaseSt = 0
                });
            }
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Rubinettos processados com sucesso.");

        return await Task.FromResult(produtos);
    }

    private static bool EhProdutoValido(string codigoFabrica, string descricao, string marca)
    {
        return !string.IsNullOrWhiteSpace(codigoFabrica)
            && !string.IsNullOrWhiteSpace(descricao)
            && !string.IsNullOrWhiteSpace(marca)
            && !codigoFabrica.Equals("REF NF", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizarMarca(string valor)
    {
        var marca = NormalizarTexto(valor);

        if (marca.Contains("RUBINETTOS COMERCIO", StringComparison.OrdinalIgnoreCase))
        {
            return "RUBINETTOS";
        }

        return marca;
    }

    private static string NormalizarUnidade(string valor)
    {
        var unidade = NormalizarTexto(valor);
        return string.IsNullOrWhiteSpace(unidade) ? "UN" : unidade;
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }

    private static string SomenteDigitos(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\D", string.Empty);
    }

    private static string CriarObservacao(string nomeAba, string cest)
    {
        var observacao = $"Importado via R3Integrador - Tabela Rubinettos - Aba {nomeAba}";
        return string.IsNullOrWhiteSpace(cest) ? observacao : $"{observacao} - CEST: {cest}";
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
