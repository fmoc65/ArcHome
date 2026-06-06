using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Text;

namespace R3Integrador.Infrastructure.Repositories;

public class DelcredereReaderService : IDelcredereReader
{
    private const string NomeAba = "COM DEL CREDERE";

    private static readonly (string Nome, int Coluna)[] TabelasPreco =
    [
        ("DEL20", 18),
        ("DEL25", 19),
        ("DEL30", 20)
    ];

    public async Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoNormalizado>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            Console.WriteLine($"Aba {NomeAba} nao encontrada.");
            return produtos;
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var estado = new LinhaDelcredere();

        for (var row = 1; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 2).GetString().Trim();

            if (string.IsNullOrWhiteSpace(referencia) || EhLinhaDeCabecalho(referencia))
            {
                continue;
            }

            estado.Atualizar(worksheet, row);

            foreach (var tabelaPreco in TabelasPreco)
            {
                var preco = ParseDecimal(worksheet.Cell(row, tabelaPreco.Coluna).GetFormattedString());

                produtos.Add(new ProdutoNormalizado
                {
                    TipoTabela = NomeAba,
                    Formato = estado.Formato,
                    Referencia = referencia,
                    Linha = estado.Linha,
                    Colecao = estado.Colecao,
                    Cor = estado.Cor,
                    Superficie = estado.Superficie,
                    Grupo = "PORCELANATO",
                    SubGrupo = NormalizarSubGrupo(estado.Superficie),
                    Marca = "VILLAGRES",
                    Modelo = estado.Formato.Replace(" ", string.Empty).ToUpper(),
                    TabelaPreco = tabelaPreco.Nome,
                    Faces = estado.Faces,
                    VariacaoTonalidade = estado.VariacaoTonalidade,
                    M2Caixa = estado.M2Caixa,
                    Espessura = estado.Espessura,
                    PrecoTabela = preco,
                    PrecoDesconto = preco,
                    PrecoVenda = CalcularPrecoVendaDelcredere(preco)
                });
            }
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} registros Del Credere processados com sucesso.");

        return await Task.FromResult(produtos);
    }

    private static bool EhLinhaDeCabecalho(string referencia)
    {
        var referenciaNormalizada = RemoverAcentos(referencia);

        return referencia.Equals("Ref.", StringComparison.OrdinalIgnoreCase)
            || referenciaNormalizada.Equals("FORMATO", StringComparison.OrdinalIgnoreCase);
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
        if (string.IsNullOrWhiteSpace(valor))
        {
            return 0;
        }

        valor = valor.Replace("R$", "")
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
        return Math.Round(preco * 1.0065m, 2);
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
        }

        private static string ManterSeVazio(string valor, string valorAtual)
        {
            valor = valor.Trim();
            return string.IsNullOrWhiteSpace(valor) ? valorAtual : valor;
        }
    }
}
