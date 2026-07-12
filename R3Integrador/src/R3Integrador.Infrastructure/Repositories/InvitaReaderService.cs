using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class InvitaReaderService : IInvitaReader
{
    private const string NomeAba = "CNPJ";
    private const string Marca = "INVITA";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        var chavesProcessadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException($"Aba '{NomeAba}' nao encontrada na tabela Invita.");
        }

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var duplicadosIgnorados = 0;

        for (var row = 92; row <= ultimaLinha; row++)
        {
            var codigo = worksheet.Cell(row, 2).GetString().Trim();
            var ncm = SomenteDigitos(worksheet.Cell(row, 6).GetString());
            var ean = worksheet.Cell(row, 7).GetString().Trim();

            if (string.IsNullOrWhiteSpace(codigo) || ncm.Length != 8)
            {
                continue;
            }

            var chave = $"{codigo}|{ean}";
            if (!chavesProcessadas.Add(chave))
            {
                duplicadosIgnorados++;
                continue;
            }

            var familia = NormalizarTexto(worksheet.Cell(row, 3).GetString());
            var descricao = NormalizarTexto(worksheet.Cell(row, 4).GetString());
            var ipi = LerPercentual(worksheet.Cell(row, 12));
            var icms = LerPercentual(worksheet.Cell(row, 14));
            var pis = LerPercentual(worksheet.Cell(row, 15));
            var cofins = LerPercentual(worksheet.Cell(row, 16));

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = codigo,
                CodigoBarras = ean,
                DescricaoCompleta = descricao,
                DescricaoComercial = descricao,
                Grupo = "ELETRODOMESTICOS E ACABAMENTOS",
                SubGrupo = familia,
                Marca = Marca,
                Linha = familia,
                Modelo = codigo,
                Voltagem = NormalizarTexto(worksheet.Cell(row, 5).GetString()),
                Ncm = ncm,
                UfOrigem = "SP",
                PrecoVenda = LerDecimal(worksheet.Cell(row, 10)),
                PrecoFabrica = LerDecimal(worksheet.Cell(row, 11)),
                DescontoPercentual = LerPercentual(worksheet.Cell(row, 9)),
                IpiPercentual = ipi,
                AliqIcmsOrigem = icms,
                AliqIcmsInterna = icms,
                Unidade = "UN",
                QtdeEmbalagemVenda = 1,
                QtdeEmbalagemCompra = 1,
                UnidFabril = "UN",
                AliquotaPisOrigem = FormatarAliquota(pis),
                AliquotaCofinsOrigem = FormatarAliquota(cofins),
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = CriarObservacao(worksheet, row, ncm),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Invita processados.");
        Console.WriteLine($"[INFO] {duplicadosIgnorados} linha(s) duplicada(s) ignorada(s) por codigo + EAN.");
        Console.WriteLine("[ATENCAO] ST/CST/CSOSN/CFOP e transicao paulista de outubro/2026 aguardam contador.");

        return Task.FromResult(produtos);
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            // Declarados por produto na tabela do fornecedor.
            [13] = SituacaoCampoFiscal.Confirmado,
            [18] = SituacaoCampoFiscal.Confirmado,
            [19] = SituacaoCampoFiscal.Confirmado,
            [20] = SituacaoCampoFiscal.Confirmado,
            [52] = SituacaoCampoFiscal.Confirmado,
            [53] = SituacaoCampoFiscal.Confirmado,

            // Dependem do enquadramento por NCM + descricao e da data da operacao.
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
            [57] = SituacaoCampoFiscal.Pendente,
            [58] = SituacaoCampoFiscal.Pendente,
            [59] = SituacaoCampoFiscal.Pendente,
            [60] = SituacaoCampoFiscal.Pendente
        };
    }

    private static string CriarObservacao(IXLWorksheet worksheet, int row, string ncm)
    {
        var observacaoOrigem = NormalizarTexto(worksheet.Cell(row, 18).GetString());
        var detalhe = string.IsNullOrWhiteSpace(observacaoOrigem) ? string.Empty : $" - {observacaoOrigem}";
        return $"Tabela Invita abril/2026 (validade encerrada em 10/05/2026){detalhe} - NCM {ncm} - IPI/ICMS/PIS/COFINS informados pela origem - ST declarada inclusa genericamente, sem CEST ou percentual por item - validar regra vigente na data da operacao";
    }

    private static decimal LerDecimal(IXLCell cell)
    {
        return cell.TryGetValue<decimal>(out var valor)
            ? valor
            : DecimalParser.Parse(cell.GetFormattedString());
    }

    private static decimal LerPercentual(IXLCell cell)
    {
        return DecimalParser.Parse(cell.GetFormattedString());
    }

    private static string FormatarAliquota(decimal valor)
    {
        return valor.ToString("0.##", System.Globalization.CultureInfo.GetCultureInfo("pt-BR"));
    }

    private static string SomenteDigitos(string valor)
    {
        return Regex.Replace(valor, @"\D", string.Empty);
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }
}
