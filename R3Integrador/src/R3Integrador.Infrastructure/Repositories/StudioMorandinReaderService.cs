using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class StudioMorandinReaderService : IStudioMorandinReader
{
    private const string NomeAba = "Plan1";
    private const string Marca = "STUDIO MORANDIN";
    private const string Ncm = "69072300";
    private const int LinhaCabecalho = 3;
    private const int PrimeiraLinhaProduto = 5;

    private static readonly string[] CabecalhosEsperados =
    [
        "LINHA",
        "REFERENCIA",
        "PRODUTO (BRICK)",
        "MEDIDA DIMENSAO CM",
        "ESPESSURA MM",
        "UNIDADE",
        "M2/CAIXA",
        "PC/M2",
        "PESO(*) CAIXA",
        "PC/CX",
        "PRECO REVENDA"
    ];

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var produtos = new List<ProdutoErpDto>();
        using var workbook = new XLWorkbook(caminhoArquivo);

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet))
        {
            throw new InvalidOperationException($"Aba '{NomeAba}' nao encontrada na tabela Studio Morandin.");
        }

        ValidarCabecalho(worksheet);

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var linhaComercial = string.Empty;

        for (var row = PrimeiraLinhaProduto; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 2).GetString().Trim();
            var produto = NormalizarTexto(worksheet.Cell(row, 3).GetString());

            if (string.IsNullOrWhiteSpace(referencia) || string.IsNullOrWhiteSpace(produto))
            {
                continue;
            }

            var linhaInformada = NormalizarTexto(worksheet.Cell(row, 1).GetString());
            if (!string.IsNullOrWhiteSpace(linhaInformada))
            {
                linhaComercial = linhaInformada;
            }

            var medida = NormalizarTexto(worksheet.Cell(row, 4).GetString());
            var unidade = NormalizarUnidade(worksheet.Cell(row, 6).GetString());
            var pesoCaixa = DecimalParser.Parse(worksheet.Cell(row, 9).GetFormattedString());
            var precoRevenda = DecimalParser.Parse(worksheet.Cell(row, 11).GetFormattedString());

            if (precoRevenda <= 0)
            {
                throw new InvalidOperationException(
                    $"Preco de revenda invalido para a referencia '{referencia}' na linha {row}.");
            }

            produtos.Add(new ProdutoErpDto
            {
                CodigoFabrica = referencia,
                DescricaoCompleta = CriarDescricao(produto, linhaComercial, medida),
                DescricaoComercial = produto,
                Grupo = "REVESTIMENTOS",
                SubGrupo = linhaComercial,
                Marca = Marca,
                Linha = linhaComercial,
                Modelo = medida,
                Ncm = Ncm,
                UfOrigem = "SP",
                // A origem chama este campo de "Preco Revenda". Ate a homologacao
                // comercial, o valor e preservado nas duas colunas sem aplicar markup.
                PrecoVenda = precoRevenda,
                PrecoFabrica = precoRevenda,
                IpiPercentual = 0,
                AliqIcmsInterna = 18,
                Iva = 81,
                Unidade = unidade,
                // M2/caixa e uma informacao logistica. A unidade comercial e o
                // metro quadrado, por isso o layout ERP recebe quantidade zero.
                QtdeEmbalagemVenda = unidade == "M2" ? 0 : 1,
                PesoBruto = pesoCaixa,
                QtdeEmbalagemCompra = 1,
                UnidFabril = unidade,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = NormalizarTexto(worksheet.Cell(row, 12).GetString()),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            });
        }

        var referenciasDuplicadas = produtos
            .GroupBy(produto => produto.CodigoFabrica, StringComparer.OrdinalIgnoreCase)
            .Where(grupo => grupo.Count() > 1)
            .Select(grupo => grupo.Key)
            .OrderBy(referencia => referencia)
            .ToList();
        if (referenciasDuplicadas.Count > 0)
        {
            throw new InvalidOperationException(
                $"Referencias duplicadas na tabela Studio Morandin: {string.Join(", ", referenciasDuplicadas)}.");
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Studio Morandin processados.");
        Console.WriteLine("[ATENCAO] Saida provisoria: CST/CSOSN, CFOP, ST efetiva, PIS/COFINS e IBS/CBS aguardam homologacao.");

        return Task.FromResult(produtos);
    }

    private static void ValidarCabecalho(IXLWorksheet worksheet)
    {
        for (var coluna = 1; coluna <= CabecalhosEsperados.Length; coluna++)
        {
            var cabecalhoEncontrado = NormalizarCabecalho(
                worksheet.Cell(LinhaCabecalho, coluna).GetString());
            var cabecalhoEsperado = NormalizarCabecalho(CabecalhosEsperados[coluna - 1]);

            if (!cabecalhoEncontrado.Equals(cabecalhoEsperado, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    $"Cabecalho Studio Morandin invalido na coluna {coluna}. " +
                    $"Esperado '{cabecalhoEsperado}', encontrado '{cabecalhoEncontrado}'.");
            }
        }
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais()
    {
        return new Dictionary<int, SituacaoCampoFiscal>
        {
            // NCM foi declarado expressamente no rodape da tabela de origem.
            [13] = SituacaoCampoFiscal.Confirmado,
            // UF de origem confirmada pelo usuario em 13/08/2026.
            [14] = SituacaoCampoFiscal.Confirmado,

            // Premissas atuais, ainda dependentes de homologacao fiscal.
            [18] = SituacaoCampoFiscal.Estimado,
            [20] = SituacaoCampoFiscal.Estimado,
            [21] = SituacaoCampoFiscal.Estimado,

            // Campos fiscais sem evidencia suficiente na origem atual.
            [19] = SituacaoCampoFiscal.Pendente,
            [26] = SituacaoCampoFiscal.Pendente,
            [27] = SituacaoCampoFiscal.Pendente,
            [28] = SituacaoCampoFiscal.Pendente,
            [29] = SituacaoCampoFiscal.Pendente,
            [30] = SituacaoCampoFiscal.Pendente,
            [31] = SituacaoCampoFiscal.Pendente,
            [32] = SituacaoCampoFiscal.Pendente,
            [37] = SituacaoCampoFiscal.Pendente,
            [38] = SituacaoCampoFiscal.Pendente,
            [39] = SituacaoCampoFiscal.Pendente,
            [51] = SituacaoCampoFiscal.Pendente,
            [52] = SituacaoCampoFiscal.Pendente,
            [53] = SituacaoCampoFiscal.Pendente,
            [57] = SituacaoCampoFiscal.Pendente,
            [58] = SituacaoCampoFiscal.Pendente,
            [59] = SituacaoCampoFiscal.Pendente,
            [60] = SituacaoCampoFiscal.Pendente
        };
    }

    private static string CriarDescricao(string produto, string linha, string medida)
    {
        return string.Join(" - ", new[] { produto, linha, medida }.Where(valor => !string.IsNullOrWhiteSpace(valor)));
    }

    private static string NormalizarUnidade(string valor)
    {
        var unidade = NormalizarCabecalho(valor);
        return unidade switch
        {
            "M2" => "M2",
            "PC" or "PCA" or "PECA" => "PC",
            _ => throw new InvalidOperationException(
                $"Unidade Studio Morandin nao reconhecida: '{valor.Trim()}'.")
        };
    }

    private static string NormalizarTexto(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ").ToUpperInvariant();
    }

    private static string NormalizarCabecalho(string valor)
    {
        var normalizado = valor.Trim().ToUpperInvariant()
            .Replace("²", "2", StringComparison.Ordinal)
            .Normalize(System.Text.NormalizationForm.FormD);
        normalizado = string.Concat(normalizado.Where(c =>
            System.Globalization.CharUnicodeInfo.GetUnicodeCategory(c) !=
            System.Globalization.UnicodeCategory.NonSpacingMark));

        return Regex.Replace(normalizado, @"[^A-Z0-9*/()]+", string.Empty);
    }
}
