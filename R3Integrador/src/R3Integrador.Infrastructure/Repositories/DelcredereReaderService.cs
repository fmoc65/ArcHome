using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public class DelcredereReaderService : IDelcredereReader
{
    private const string NomeAba = "COM DEL CREDERE";
    private const decimal IpiPercentual = 0.65m;
    private const decimal TaxaCartaoPercentual = 4.71m;

    private static readonly (string Nome, int Coluna)[] TabelasPreco =
    [
        ("DEL5", 18),
        ("DEL10", 19),
        ("DEL15", 20),
        ("DEL20", 21),
        ("DEL25", 22),
        ("DEL30", 23)
    ];

    public async Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo, string? caminhoAliquotas = null)
    {
        var produtos = new List<ProdutoNormalizado>();
        var aliquotas = string.IsNullOrWhiteSpace(caminhoAliquotas)
            ? null
            : await CarregarAliquotasAsync(caminhoAliquotas);

        using var workbook = new XLWorkbook(caminhoArquivo);

        var worksheet = ObterPlanilhaDelcredere(workbook);
        if (worksheet == null)
        {
            Console.WriteLine($"Aba Del Credere nao encontrada. Esperada: {NomeAba}.");
            return produtos;
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var estado = new LinhaDelcredere();
        var legendaSegmentos = CarregarLegendaSegmentos(worksheet);

        for (var row = 1; row <= ultimaLinha; row++)
        {
            var referenciaOriginal = worksheet.Cell(row, 2).GetString().Trim();

            if (!EhReferenciaValida(referenciaOriginal))
            {
                continue;
            }

            var referencia = ExtrairReferencia(referenciaOriginal);
            estado.Atualizar(worksheet, row);
            var segmentacao = DeterminarSegmentacao(worksheet, row, legendaSegmentos);

            foreach (var tabelaPreco in TabelasPreco)
            {
                var preco = ParseDecimal(worksheet.Cell(row, tabelaPreco.Coluna).GetFormattedString());
                var produto = new ProdutoNormalizado
                {
                    TipoTabela = NomeAba,
                    Formato = estado.Formato,
                    Referencia = referencia,
                    Linha = estado.Linha,
                    Colecao = estado.Colecao,
                    Cor = estado.Cor,
                    SegmentacaoComercial = segmentacao,
                    Superficie = estado.Superficie,
                    Grupo = "PORCELANATO",
                    SubGrupo = NormalizarSubGrupo(estado.Superficie),
                    Marca = DeterminarMarca(referenciaOriginal),
                    Modelo = estado.Formato.Replace(" ", string.Empty).ToUpper(),
                    TabelaPreco = tabelaPreco.Nome,
                    Faces = estado.Faces,
                    VariacaoTonalidade = estado.VariacaoTonalidade,
                    M2Caixa = estado.M2Caixa,
                    Espessura = estado.Espessura,
                    PesoBrutoM2 = estado.PesoBrutoM2,
                    PrecoTabela = preco,
                    PrecoDesconto = preco,
                    PrecoVenda = CalcularPrecoVendaDelcredere(preco)
                };

                if (aliquotas != null && aliquotas.TryGetValue(referencia, out var aliquota))
                {
                    produto.AliquotaIbs = aliquota.AliquotaIbs;
                    produto.AliquotaCbs = aliquota.AliquotaCbs;
                    produto.ClassificacaoTributaria = aliquota.ClassificacaoTributaria;
                    produto.CodigoBeneficio = aliquota.CodigoBeneficio;
                }

                produtos.Add(produto);
            }
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} registros Del Credere processados com sucesso.");
        var segmentos = produtos
            .GroupBy(produto => string.IsNullOrWhiteSpace(produto.SegmentacaoComercial)
                ? "SEM SEGMENTACAO VISUAL"
                : produto.SegmentacaoComercial)
            .OrderBy(grupo => grupo.Key)
            .Select(grupo => $"{grupo.Key}: {grupo.Select(produto => produto.Referencia).Distinct().Count()}");
        Console.WriteLine($"[INFO] Segmentacao visual B/C: {string.Join("; ", segmentos)}.");

        return await Task.FromResult(produtos);
    }

    private static IXLWorksheet? ObterPlanilhaDelcredere(XLWorkbook workbook)
    {
        if (workbook.TryGetWorksheet(NomeAba, out var worksheetPadrao))
        {
            return worksheetPadrao;
        }

        // A tabela atual recebida foi nomeada "100 TABELA COM DEL CREDERE".
        // Aceitar variações do título, validando também o cabeçalho das seis faixas.
        return workbook.Worksheets.FirstOrDefault(worksheet =>
            worksheet.Name.Contains("DEL CREDERE", StringComparison.OrdinalIgnoreCase)
            && worksheet.Cell(3, 2).GetString().Trim()
                .Equals("REF.", StringComparison.OrdinalIgnoreCase)
            && worksheet.Cell(4, 18).GetString().Contains("DEL", StringComparison.OrdinalIgnoreCase)
            && worksheet.Cell(4, 23).GetString().Contains("DEL", StringComparison.OrdinalIgnoreCase));
    }

    private static IReadOnlyCollection<SegmentoLegenda> CarregarLegendaSegmentos(IXLWorksheet worksheet)
    {
        var segmentos = new List<SegmentoLegenda>();
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 1; row <= ultimaLinha; row++)
        {
            var descricao = worksheet.Cell(row, 2).GetString().Trim();
            var nome = descricao switch
            {
                "VILLA PREMIUM" => "VILLA PREMIUM",
                "VILLA EXCLUSIVE" => "VILLA EXCLUSIVE",
                "VILLA MAX" => "VILLA MAX",
                "VILLA STYLE (e também os itens **)" => "VILLA STYLE",
                "EXCLUSIVO VILLA ART" => "EXCLUSIVO VILLA ART",
                _ => string.Empty
            };

            if (!string.IsNullOrWhiteSpace(nome))
            {
                segmentos.Add(new SegmentoLegenda(nome, worksheet.Cell(row, 2).Style.Fill.BackgroundColor));
            }
        }

        return segmentos;
    }

    private static string DeterminarSegmentacao(
        IXLWorksheet worksheet,
        int row,
        IReadOnlyCollection<SegmentoLegenda> legendaSegmentos)
    {
        var corReferencia = worksheet.Cell(row, 2).Style.Fill.BackgroundColor;
        var corLinha = worksheet.Cell(row, 3).Style.Fill.BackgroundColor;
        var segmento = legendaSegmentos.FirstOrDefault(item =>
            item.Cor.Equals(corReferencia) || item.Cor.Equals(corLinha));

        return segmento?.Nome ?? string.Empty;
    }

    private static async Task<Dictionary<string, AliquotaLinha>> CarregarAliquotasAsync(string caminhoAliquotas)
    {
        var lookup = new Dictionary<string, AliquotaLinha>(StringComparer.OrdinalIgnoreCase);
        using var workbook = new XLWorkbook(caminhoAliquotas);
        var worksheet = workbook.Worksheets.FirstOrDefault() ?? workbook.Worksheet(1);
        if (worksheet == null)
        {
            return lookup;
        }

        var headerCells = worksheet.Row(1).CellsUsed().ToDictionary(
            c => RemoverAcentos(c.GetString()).Trim().ToUpperInvariant(),
            c => c.Address.ColumnNumber);

        var codigoFabricaCol = ObterColuna(headerCells, "CÓDIGO FÁBRICA", "CODIGO FABRICA", "B");
        var aliquotaIbsCol = ObterColuna(headerCells, "ALIQUOTA IBS", "BE");
        var aliquotaCbsCol = ObterColuna(headerCells, "ALIQUOTA CBS", "BF");
        var classificacaoTributariaCol = ObterColuna(headerCells, "CLASSIFICACAO TRIBUTARIA", "BG");
        var codigoBeneficioCol = ObterColuna(headerCells, "CODIGO BENEFICIO", "BH");

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        for (var row = 2; row <= ultimaLinha; row++)
        {
            var codigo = worksheet.Cell(row, codigoFabricaCol).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(codigo))
            {
                continue;
            }

            lookup[codigo] = new AliquotaLinha
            {
                AliquotaIbs = GetCellString(worksheet, row, aliquotaIbsCol),
                AliquotaCbs = GetCellString(worksheet, row, aliquotaCbsCol),
                ClassificacaoTributaria = GetCellString(worksheet, row, classificacaoTributariaCol),
                CodigoBeneficio = GetCellString(worksheet, row, codigoBeneficioCol)
            };
        }

        return await Task.FromResult(lookup);
    }

    private static string GetCellString(IXLWorksheet worksheet, int row, int column)
    {
        return column <= 0 ? string.Empty : worksheet.Cell(row, column).GetFormattedString().Trim();
    }

    private static int ObterColuna(Dictionary<string, int> headerCells, params string[] nomes)
    {
        foreach (var nome in nomes)
        {
            var chave = RemoverAcentos(nome).Trim().ToUpperInvariant();
            if (headerCells.TryGetValue(chave, out var coluna))
            {
                return coluna;
            }
        }

        return 0;
    }

    private static bool EhLinhaDeCabecalho(string referencia)
    {
        var referenciaNormalizada = RemoverAcentos(referencia);

        return referencia.Equals("Ref.", StringComparison.OrdinalIgnoreCase)
            || referenciaNormalizada.Equals("FORMATO", StringComparison.OrdinalIgnoreCase);
    }

    private static bool EhReferenciaValida(string referencia)
    {
        return !string.IsNullOrWhiteSpace(referencia) &&
            !EhLinhaDeCabecalho(referencia) &&
            Regex.IsMatch(referencia, @"^\d+");
    }

    private static string ExtrairReferencia(string referencia)
    {
        var match = Regex.Match(referencia, @"^\d+");
        return match.Success ? match.Value : referencia.Trim();
    }

    private static string DeterminarMarca(string referencia)
    {
        return referencia.Contains("VILLA ART", StringComparison.OrdinalIgnoreCase)
            ? "VILLA ART"
            : "VILLAGRES";
    }

    private static string RemoverAcentos(string valor)
    {
        var normalizado = valor.Normalize(NormalizationForm.FormD);
        var builder = new StringBuilder();

        foreach (var caractere in normalizado)
        {
            if (CharUnicodeInfo.GetUnicodeCategory(caractere) != UnicodeCategory.NonSpacingMark)
            {
                builder.Append(caractere);
            }
        }

        return builder.ToString().Normalize(NormalizationForm.FormC);
    }

    private static decimal ParseDecimal(string valor)
    {
        return DecimalParser.Parse(valor);
    }

    private static int ParseInt(string valor)
    {
        int.TryParse(valor, out var resultado);
        return resultado;
    }

    private static string NormalizarSubGrupo(string superficie)
    {
        return superficie.Equals("NATURAL SENSE UP", StringComparison.OrdinalIgnoreCase)
            ? "NATURAL SENSEUP"
            : superficie.ToUpper();
    }

    private static decimal CalcularPrecoVendaDelcredere(decimal preco)
    {
        var acrescimoPercentual = (IpiPercentual + TaxaCartaoPercentual) / 100m;
        return Math.Round(preco * (1m + acrescimoPercentual), 2);
    }

    private sealed class LinhaDelcredere
    {
        public string Formato { get; private set; } = string.Empty;
        public string Linha { get; private set; } = string.Empty;
        public string Colecao { get; private set; } = string.Empty;
        public string Cor { get; private set; } = string.Empty;
        public string Superficie { get; private set; } = string.Empty;
        public int Faces { get; private set; }
        public string VariacaoTonalidade { get; private set; } = string.Empty;
        public decimal M2Caixa { get; private set; }
        public decimal Espessura { get; private set; }
        public decimal PesoBrutoM2 { get; private set; }

        public void Atualizar(IXLWorksheet worksheet, int row)
        {
            Formato = ManterSeVazio(worksheet.Cell(row, 1).GetString(), Formato);
            Linha = ManterSeVazio(worksheet.Cell(row, 3).GetString(), Linha);
            Colecao = ManterSeVazio(worksheet.Cell(row, 4).GetString(), Colecao);
            Cor = ManterSeVazio(worksheet.Cell(row, 5).GetString(), Cor);
            Superficie = ManterSeVazio(worksheet.Cell(row, 6).GetString(), Superficie);
            VariacaoTonalidade = ManterSeVazio(worksheet.Cell(row, 8).GetString(), VariacaoTonalidade);

            var faces = worksheet.Cell(row, 7).GetString();
            if (!string.IsNullOrWhiteSpace(faces))
            {
                Faces = ParseInt(faces);
            }

            var m2Caixa = worksheet.Cell(row, 11).GetString();
            if (!string.IsNullOrWhiteSpace(m2Caixa))
            {
                M2Caixa = ParseDecimal(m2Caixa);
            }

            var espessura = worksheet.Cell(row, 17).GetString();
            if (!string.IsNullOrWhiteSpace(espessura))
            {
                Espessura = ParseDecimal(espessura);
            }

            var pesoBruto = worksheet.Cell(row, 15).GetString();
            if (!string.IsNullOrWhiteSpace(pesoBruto))
            {
                PesoBrutoM2 = ParseDecimal(pesoBruto);
            }
        }

        private static string ManterSeVazio(string valor, string valorAtual)
        {
            valor = valor.Trim();
            return string.IsNullOrWhiteSpace(valor) ? valorAtual : valor;
        }
    }

    private sealed class AliquotaLinha
    {
        public string AliquotaIbs { get; init; } = string.Empty;
        public string AliquotaCbs { get; init; } = string.Empty;
        public string ClassificacaoTributaria { get; init; } = string.Empty;
        public string CodigoBeneficio { get; init; } = string.Empty;
    }

    private sealed record SegmentoLegenda(string Nome, XLColor Cor);
}
