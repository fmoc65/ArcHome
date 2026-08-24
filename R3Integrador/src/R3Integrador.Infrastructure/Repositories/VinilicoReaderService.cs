using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Globalization;
using System.Text;

namespace R3Integrador.Infrastructure.Repositories;

public class VinilicoReaderService : IVinilicoReader
{
    private const string NomeAba = "VINILICO";
    private const string NomeAbaAlternativa = "VAREJO";

    public async Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoNormalizado>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet) &&
            !workbook.TryGetWorksheet(NomeAbaAlternativa, out worksheet))
        {
            Console.WriteLine($"Abas {NomeAba} e {NomeAbaAlternativa} nao encontradas.");
            return produtos;
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var estado = new LinhaVinilico();

        for (var row = 1; row <= ultimaLinha; row++)
        {
            var referencia = NormalizarReferencia(worksheet.Cell(row, 2).GetString());

            if (string.IsNullOrWhiteSpace(referencia) || EhLinhaDeCabecalho(referencia))
            {
                continue;
            }

            estado.Atualizar(worksheet, row);
            var precoDesconto = ParseDecimal(worksheet.Cell(row, 19).GetFormattedString());
            var precoFinal = CalcularPrecoFinal(precoDesconto);

            produtos.Add(new ProdutoNormalizado
            {
                TipoTabela = NomeAba,
                Formato = estado.Formato,
                Referencia = referencia,
                Linha = estado.Linha,
                Colecao = estado.Colecao,
                Cor = estado.Cor,
                Superficie = estado.Superficie,
                Grupo = "VINILICO",
                SubGrupo = DeterminarSubGrupo(referencia, estado.Superficie),
                Marca = "VILLAGRES",
                Modelo = estado.Formato.Replace(" ", string.Empty).ToUpper(),
                Faces = estado.Faces,
                VariacaoTonalidade = estado.CapaDesgaste,
                M2Caixa = estado.M2Caixa,
                Espessura = estado.Espessura,
                PesoBrutoM2 = estado.PesoBrutoM2,
                PrecoTabela = precoFinal,
                PrecoDesconto = precoDesconto,
                PrecoVenda = precoFinal
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos vinilicos processados com sucesso.");

        return await Task.FromResult(produtos);
    }

    private static bool EhLinhaDeCabecalho(string referencia)
    {
        var referenciaNormalizada = RemoverAcentos(referencia);

        return referencia.Equals("Ref.", StringComparison.OrdinalIgnoreCase)
            || referenciaNormalizada.Equals("VINILICOS", StringComparison.OrdinalIgnoreCase)
            || referencia.Equals("Formato", StringComparison.OrdinalIgnoreCase);
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

    private static string NormalizarReferencia(string valor)
    {
        return valor.Replace("***", string.Empty, StringComparison.Ordinal).Trim();
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

    private static string DeterminarSubGrupo(string referencia, string superficie)
    {
        if (referencia.StartsWith("SPC", StringComparison.OrdinalIgnoreCase))
        {
            return "CLICADO";
        }

        if (referencia.StartsWith("LVT", StringComparison.OrdinalIgnoreCase))
        {
            return "COLADO";
        }

        return superficie.ToUpper();
    }

    private static decimal CalcularPrecoFinal(decimal precoDesconto)
    {
        const decimal percentualStVinilico = 7.92m;
        var fator = 1m + percentualStVinilico / 100m;
        return Math.Round(precoDesconto * fator, 2, MidpointRounding.AwayFromZero);
    }

    private sealed class LinhaVinilico
    {
        public string Formato { get; private set; } = string.Empty;
        public string Linha { get; private set; } = string.Empty;
        public string Colecao { get; private set; } = string.Empty;
        public string Cor { get; private set; } = string.Empty;
        public string Superficie { get; private set; } = string.Empty;
        public int Faces { get; private set; }
        public string CapaDesgaste { get; private set; } = string.Empty;
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
            CapaDesgaste = ManterSeVazio(worksheet.Cell(row, 8).GetString(), CapaDesgaste);

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
}
