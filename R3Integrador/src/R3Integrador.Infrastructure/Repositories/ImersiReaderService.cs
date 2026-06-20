using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public class ImersiReaderService : IImersiReader
{
    private const string NomeAba = "LISTA DE PREÇOS";
    private const string Marca = "IMERSI";

    private static readonly (string Nome, int Coluna)[] TabelasPreco =
    [
        ("7X_A_10X", 7),
        ("4X_A_6X", 8),
        ("ATE_3X", 9),
        ("ANTECIPADO", 10)
    ];

    public async Task<Dictionary<string, List<ProdutoErpDto>>> LerAsync(string caminhoArquivo)
    {
        var produtosPorTabela = TabelasPreco.ToDictionary(tabela => tabela.Nome, _ => new List<ProdutoErpDto>());
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            Console.WriteLine($"Aba {NomeAba} nao encontrada.");
            return produtosPorTabela;
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 6; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 1).GetString().Trim();
            var descricao = worksheet.Cell(row, 4).GetString().Trim();
            var ncm = SomenteDigitos(worksheet.Cell(row, 6).GetString());

            if (!EhProdutoValido(referencia, descricao, ncm))
            {
                continue;
            }

            foreach (var tabelaPreco in TabelasPreco)
            {
                var preco = ParseDecimal(worksheet.Cell(row, tabelaPreco.Coluna).GetFormattedString());

                if (preco <= 0)
                {
                    continue;
                }

                produtosPorTabela[tabelaPreco.Nome].Add(MapearProduto(worksheet, row, referencia, descricao, ncm, preco, tabelaPreco.Nome));
            }
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtosPorTabela.Sum(tabela => tabela.Value.Count)} registros Imersi processados com sucesso.");

        return await Task.FromResult(produtosPorTabela);
    }

    private static ProdutoErpDto MapearProduto(
        IXLWorksheet worksheet,
        int row,
        string referencia,
        string descricao,
        string ncm,
        decimal preco,
        string tabelaPreco)
    {
        var linha = NormalizarTexto(worksheet.Cell(row, 3).GetString());
        var tamanho = worksheet.Cell(row, 5).GetString().Trim();
        var tributacao = ObterTributacao(ncm);

        return new ProdutoErpDto
        {
            CodigoInterno = string.Empty,
            CodigoFabrica = referencia,
            CodigoBarras = string.Empty,
            DescricaoCompleta = NormalizarTexto(descricao),
            DescricaoComercial = NormalizarTexto(descricao),
            Grupo = ObterGrupo(ncm),
            SubGrupo = string.IsNullOrWhiteSpace(linha) ? ObterGrupo(ncm) : linha,
            Marca = Marca,
            Linha = linha,
            Modelo = tamanho,
            Voltagem = string.Empty,
            Cor = ObterCor(descricao),
            Ncm = ncm,
            UfOrigem = "SP",
            PrecoVenda = preco,
            PrecoFabrica = preco,
            DescontoPercentual = 0,
            IpiPercentual = 0,
            AliqIcmsOrigem = 18,
            AliqIcmsInterna = 18,
            Iva = tributacao.MvaOriginal,
            FreteReais = 0,
            FretePercentual = 0,
            Unidade = "UN",
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
            PercentualSt = tributacao.MvaOriginal,
            UnidFabril = "UN",
            Observacao = CriarObservacao(tabelaPreco, tamanho, tributacao),
            DiferencaIcms = 0,
            ReducaoBaseIcms = 0,
            ReducaoBaseSt = 0
        };
    }

    private static bool EhProdutoValido(string referencia, string descricao, string ncm)
    {
        return !string.IsNullOrWhiteSpace(referencia)
            && !string.IsNullOrWhiteSpace(descricao)
            && !string.IsNullOrWhiteSpace(ncm)
            && !referencia.Equals("REF.", StringComparison.OrdinalIgnoreCase)
            && !referencia.StartsWith("OBSERVAÇÃO", StringComparison.OrdinalIgnoreCase);
    }

    private static string ObterGrupo(string ncm)
    {
        return ncm switch
        {
            "39221000" => "BANHEIRAS",
            "84818019" => "METAIS",
            _ => "BANHO"
        };
    }

    private static TributacaoImersi ObterTributacao(string ncm)
    {
        return ncm switch
        {
            "39221000" => new TributacaoImersi("1001300", 91m, 123.61m, 104.98m),
            "84818019" => new TributacaoImersi("1007900", 85m, 116.59m, 98.54m),
            _ => new TributacaoImersi(string.Empty, 0, 0, 0)
        };
    }

    private static string ObterCor(string descricao)
    {
        var descricaoNormalizada = NormalizarTexto(descricao);

        var cores = new[]
        {
            "GLOSSY WHITE",
            "MATT WHITE",
            "PURE BLACK",
            "DEEP BLUE",
            "FRESH GREEN",
            "SOFT TERRACOTTA",
            "TAN WHITE",
            "WARM GREY",
            "CHROME",
            "BRUSHED NICKEL",
            "GOLD",
            "MATT BLACK",
            "ROSE GOLD",
            "BRANCOS",
            "CROMADOS"
        };

        return cores.FirstOrDefault(cor => descricaoNormalizada.Contains(cor, StringComparison.OrdinalIgnoreCase)) ?? string.Empty;
    }

    private static string CriarObservacao(string tabelaPreco, string tamanho, TributacaoImersi tributacao)
    {
        var observacao = $"Importado via R3Integrador - Tabela Imersi - Preco {tabelaPreco}";

        if (!string.IsNullOrWhiteSpace(tamanho))
        {
            observacao += $" - Tamanho: {tamanho}";
        }

        if (!string.IsNullOrWhiteSpace(tributacao.Cest))
        {
            observacao += $" - CEST: {tributacao.Cest}";
        }

        if (tributacao.MvaOriginal > 0)
        {
            observacao += $" - MVA original: {tributacao.MvaOriginal:N2}%";
            observacao += $" - MVA ajustada 4%: {tributacao.MvaAjustada4:N2}%";
            observacao += $" - MVA ajustada 12%: {tributacao.MvaAjustada12:N2}%";
        }

        return observacao;
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

    private sealed record TributacaoImersi(
        string Cest,
        decimal MvaOriginal,
        decimal MvaAjustada4,
        decimal MvaAjustada12);
}
