using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class NinaMartinelliReaderService : INinaMartinelliReader
{
    private const string NomeAba = "COLECAO_COMPLETA";
    private const string NomeAbaImportacaoErp = "IMPORTACAO_ERP";
    private const string Marca = "NINA MARTINELLI";

    // Regras recebidas do contador em 16/07/2026. A ST permanece estimada ate o
    // recebimento da relacao de produtos sem incidencia.
    private const decimal IpiPercentualContador = 0.65m;
    private const decimal AliqIcmsOrigemContador = 12m;
    private const decimal AliqIcmsInternaContador = 12m;
    private const decimal IvaContador = 81m;
    private const decimal PercentualStPadraoContador = 9.86m;
    private const string AliquotaPisOrigemContador = "1,65";
    private const string AliquotaCofinsOrigemContador = "7,60";

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        using var workbook = new XLWorkbook(caminhoArquivo);
        if (workbook.TryGetWorksheet(NomeAbaImportacaoErp, out var worksheetImportacaoErp))
        {
            return Task.FromResult(LerImportacaoErp(worksheetImportacaoErp));
        }

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

            var produto = new ProdutoErpDto
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
            };
            AplicarTributacaoContador(produto);
            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Nina Martinelli processados.");
        Console.WriteLine($"[INFO] {referenciasTecnicasCriadas} referencia(s) tecnica(s) criada(s) para codigos repetidos.");
        Console.WriteLine("[ATENCAO] NCM e demais campos fiscais sem orientacao do contador continuam pendentes.");
        return Task.FromResult(produtos);
    }

    private static List<ProdutoErpDto> LerImportacaoErp(IXLWorksheet worksheet)
    {
        var produtos = new List<ProdutoErpDto>();
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 2; row <= ultimaLinha; row++)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Cell(row, 2).GetFormattedString()))
            {
                continue;
            }

            var produto = new ProdutoErpDto
            {
                CodigoInterno = Texto(worksheet, row, 1), CodigoFabrica = Texto(worksheet, row, 2), CodigoBarras = Texto(worksheet, row, 3),
                DescricaoCompleta = Texto(worksheet, row, 4), DescricaoComercial = Texto(worksheet, row, 5), Grupo = Texto(worksheet, row, 6),
                SubGrupo = Texto(worksheet, row, 7), Marca = Texto(worksheet, row, 8), Linha = Texto(worksheet, row, 9), Modelo = Texto(worksheet, row, 10),
                Voltagem = Texto(worksheet, row, 11), Cor = Texto(worksheet, row, 12), Ncm = Texto(worksheet, row, 13), UfOrigem = Texto(worksheet, row, 14),
                PrecoVenda = Decimal(worksheet, row, 15), PrecoFabrica = Decimal(worksheet, row, 16), DescontoPercentual = Decimal(worksheet, row, 17),
                IpiPercentual = Decimal(worksheet, row, 18), AliqIcmsOrigem = Decimal(worksheet, row, 19), AliqIcmsInterna = Decimal(worksheet, row, 20),
                Iva = Decimal(worksheet, row, 21), FreteReais = Decimal(worksheet, row, 22), FretePercentual = Decimal(worksheet, row, 23), Unidade = Texto(worksheet, row, 24),
                QtdeEmbalagemVenda = Decimal(worksheet, row, 25), Cst = Texto(worksheet, row, 26), AliquotaCofinsCst = Texto(worksheet, row, 27),
                AliquotaIpiCst = Texto(worksheet, row, 28), AliquotaPisCst = Texto(worksheet, row, 29), Csosn = Texto(worksheet, row, 30),
                CfopDentro = Texto(worksheet, row, 31), CfopFora = Texto(worksheet, row, 32), PesoLiquido = Decimal(worksheet, row, 33), PesoBruto = Decimal(worksheet, row, 34),
                QtdeEmbalagemCompra = Decimal(worksheet, row, 35), ValorPi = Decimal(worksheet, row, 36), AliquotaCofins = Decimal(worksheet, row, 37),
                AliquotaPis = Decimal(worksheet, row, 38), PercentualSt = Decimal(worksheet, row, 39), UnidFabril = Texto(worksheet, row, 40),
                Observacao = Texto(worksheet, row, 41), DiferencaIcms = Decimal(worksheet, row, 42), ReducaoBaseIcms = Decimal(worksheet, row, 43),
                ReducaoBaseSt = Decimal(worksheet, row, 44), RetencaoPis = Texto(worksheet, row, 45), RetencaoCofins = Texto(worksheet, row, 46),
                RetencaoCsll = Texto(worksheet, row, 47), RetencaoIrrf = Texto(worksheet, row, 48), RetencaoPrevSocial = Texto(worksheet, row, 49),
                Localizacao = Texto(worksheet, row, 50), EnquadramentoIpi = Texto(worksheet, row, 51), AliquotaPisOrigem = Texto(worksheet, row, 52),
                AliquotaCofinsOrigem = Texto(worksheet, row, 53), Imagem = Texto(worksheet, row, 54), EstoqueMinimo = Decimal(worksheet, row, 55),
                EstoqueMaximo = Decimal(worksheet, row, 56), AliquotaIbs = Texto(worksheet, row, 57), AliquotaCbs = Texto(worksheet, row, 58),
                ClassificacaoTributaria = Texto(worksheet, row, 59), CodigoBeneficio = Texto(worksheet, row, 60),
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            };
            AplicarTributacaoContador(produto);
            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Nina Martinelli atualizados a partir de uma importacao ERP existente.");
        return produtos;
    }

    private static void AplicarTributacaoContador(ProdutoErpDto produto)
    {
        produto.IpiPercentual = IpiPercentualContador;
        produto.AliqIcmsOrigem = AliqIcmsOrigemContador;
        produto.AliqIcmsInterna = AliqIcmsInternaContador;
        produto.Iva = IvaContador;
        produto.PercentualSt = PercentualStPadraoContador;
        produto.AliquotaPisOrigem = AliquotaPisOrigemContador;
        produto.AliquotaCofinsOrigem = AliquotaCofinsOrigemContador;

        foreach (var coluna in new[] { 18, 19, 20, 21, 52, 53 })
        {
            produto.SituacaoCamposFiscais[coluna] = SituacaoCampoFiscal.Confirmado;
        }

        produto.SituacaoCamposFiscais[39] = SituacaoCampoFiscal.Estimado;
    }

    private static string Texto(IXLWorksheet worksheet, int row, int column) => worksheet.Cell(row, column).GetFormattedString().Trim();

    private static decimal Decimal(IXLWorksheet worksheet, int row, int column) => LerDecimal(worksheet.Cell(row, column));

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
