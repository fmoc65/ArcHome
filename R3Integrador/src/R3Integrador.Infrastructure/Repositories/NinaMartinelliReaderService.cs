using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class NinaMartinelliReaderService : INinaMartinelliReader
{
    private const string NomeAba = "COLECAO_COMPLETA";
    private const string Marca = "NINA MARTINELLI";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        using var workbook = new XLWorkbook(caminhoArquivo);
        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException($"Aba '{NomeAba}' nao encontrada na tabela Nina Martinelli.");
        }

        var produtos = new List<ProdutoErpDto>();
        var referenciasUsadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var referenciasTecnicasCriadas = 0;
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 2; row <= ultimaLinha; row++)
        {
            var nome = NormalizarTexto(worksheet.Cell(row, 4).GetString());
            var codigoOrigem = worksheet.Cell(row, 3).GetFormattedString().Trim();
            if (string.IsNullOrWhiteSpace(nome) || string.IsNullOrWhiteSpace(codigoOrigem))
            {
                continue;
            }

            var aplicacao = NormalizarTexto(worksheet.Cell(row, 5).GetString());
            var cor = NormalizarTexto(worksheet.Cell(row, 6).GetString());
            var dimensao = NormalizarTexto(worksheet.Cell(row, 7).GetString());
            var espessura = NormalizarTexto(worksheet.Cell(row, 8).GetString());
            var unidade = NormalizarUnidade(worksheet.Cell(row, 18).GetString());
            var embalagem = LerDecimal(worksheet.Cell(row, 11));
            var preco = LerDecimal(worksheet.Cell(row, 13));

            var codigo = codigoOrigem;
            if (codigoOrigem.Equals("depende do raio", StringComparison.OrdinalIgnoreCase) || !referenciasUsadas.Add(codigo))
            {
                // "depende do raio" e uma descricao de referencia, nao uma chave ERP.
                // A chave derivada e deterministica e distingue cada borda pelo catalogo.
                codigo = CriarReferenciaTecnica(nome, cor, dimensao, row);
                referenciasUsadas.Add(codigo);
                referenciasTecnicasCriadas++;
            }

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = codigo,
                DescricaoCompleta = CriarDescricao(nome, cor, dimensao, espessura),
                DescricaoComercial = nome,
                Grupo = "REVESTIMENTOS",
                SubGrupo = string.IsNullOrWhiteSpace(aplicacao) ? "REVESTIMENTOS" : aplicacao,
                Marca = Marca,
                Linha = aplicacao,
                Modelo = dimensao,
                Cor = cor,
                UfOrigem = NormalizarTexto(worksheet.Cell(row, 19).GetString()),
                // O preco de representacao e preservado provisoriamente nas duas colunas;
                // frete, margem e tributacao ainda nao foram homologados.
                PrecoVenda = preco,
                PrecoFabrica = preco,
                Unidade = unidade,
                QtdeEmbalagemVenda = CalcularEmbalagemVenda(unidade, embalagem),
                PesoBruto = LerDecimal(worksheet.Cell(row, 14)),
                QtdeEmbalagemCompra = 1,
                UnidFabril = unidade,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = CriarObservacao(worksheet, row, codigoOrigem, codigo, embalagem),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Nina Martinelli processados.");
        Console.WriteLine($"[INFO] {referenciasTecnicasCriadas} referencia(s) tecnica(s) criada(s) para codigos repetidos.");
        Console.WriteLine("[ATENCAO] NCM e tributacao nao constam na origem; arquivo provisório aguarda fornecedor/contador.");
        return Task.FromResult(produtos);
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais() => new()
    {
        [13] = SituacaoCampoFiscal.Pendente,
        [18] = SituacaoCampoFiscal.Pendente,
        [19] = SituacaoCampoFiscal.Pendente,
        [20] = SituacaoCampoFiscal.Pendente,
        [21] = SituacaoCampoFiscal.Pendente,
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

    private static string CriarObservacao(IXLWorksheet worksheet, int row, string codigoOrigem, string codigo, decimal embalagem)
    {
        var pecasM2 = worksheet.Cell(row, 9).GetFormattedString().Trim();
        var pecasCaixa = worksheet.Cell(row, 10).GetFormattedString().Trim();
        var unidadeEmbalagem = NormalizarTexto(worksheet.Cell(row, 12).GetString());
        var acabamento = NormalizarTexto(worksheet.Cell(row, 21).GetString());
        var adicional = NormalizarTexto(worksheet.Cell(row, 22).GetString());
        var codigoTecnico = codigo == codigoOrigem ? string.Empty : $" - codigo origem: {codigoOrigem}";
        return $"Nina Martinelli representacao REV02 - embalagem: {embalagem:0.####} {unidadeEmbalagem} - {pecasM2} - {pecasCaixa} - acabamento: {acabamento} - {adicional}{codigoTecnico} - NCM, fiscal e regra comercial pendentes de homologacao";
    }

    private static decimal CalcularEmbalagemVenda(string unidade, decimal embalagem) =>
        unidade == "M2" && embalagem > 0 ? embalagem : 1;

    private static string CriarReferenciaTecnica(string nome, string cor, string dimensao, int row)
    {
        var baseReferencia = Regex.Replace($"BORDA-{nome}-{cor}-{dimensao}", @"[^A-Z0-9]+", "-").Trim('-');
        return $"{baseReferencia}-{row}";
    }

    private static string CriarDescricao(params string[] partes) =>
        string.Join(" - ", partes.Where(parte => !string.IsNullOrWhiteSpace(parte)));

    private static decimal LerDecimal(IXLCell cell) =>
        cell.TryGetValue<decimal>(out var valor) ? valor : DecimalParser.Parse(cell.GetFormattedString());

    private static string NormalizarUnidade(string valor)
    {
        var unidade = NormalizarTexto(valor);
        if (unidade.Contains("M²") || unidade == "M2") return "M2";
        if (unidade.Contains("LITRO") || unidade == "L") return "LT";
        return "PC";
    }

    private static string NormalizarTexto(string valor) => Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
}
