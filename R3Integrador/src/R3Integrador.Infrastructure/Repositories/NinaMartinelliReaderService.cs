using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Services;
using System.Text.RegularExpressions;

namespace R3Integrador.Infrastructure.Repositories;

public sealed class NinaMartinelliReaderService : INinaMartinelliReader
{
    private const string NomeAba = "COLECAO_COMPLETA";
    private const string NomeAbaRev04 = "Coleção Completa";
    private const string NomeAbaImportacaoErp = "IMPORTACAO_ERP";
    private const string Marca = "NINA MARTINELLI";
    private static readonly string[] PrefixosProdutosCimenticios = ["154", "158", "159"];

    private readonly ICsosnService _csosnService;
    private readonly NinaFiscalParameters _fiscalParameters;

    public NinaMartinelliReaderService(
        ICsosnService csosnService,
        NinaFiscalParameters fiscalParameters)
    {
        _csosnService = csosnService;
        _fiscalParameters = fiscalParameters;
    }

    public Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        using var workbook = new XLWorkbook(caminhoArquivo);
        if (workbook.TryGetWorksheet(NomeAbaImportacaoErp, out var worksheetImportacaoErp))
        {
            return Task.FromResult(LerImportacaoErp(worksheetImportacaoErp));
        }

        if (!workbook.TryGetWorksheet(NomeAba, out var worksheet) &&
            !workbook.TryGetWorksheet(NomeAbaRev04, out worksheet))
        {
            throw new InvalidOperationException(
                $"Aba '{NomeAba}' ou '{NomeAbaRev04}' nao encontrada na tabela Nina Martinelli.");
        }

        var produtos = new List<ProdutoErpDto>();
        var referenciasUsadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var referenciasTecnicasCriadas = 0;
        var codigosFabricaAusentes = 0;
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
            var tributacaoSt = LerTributacaoSt(worksheet, row);

            var codigo = codigoOrigem;
            if (codigoOrigem.Equals("depende do raio", StringComparison.OrdinalIgnoreCase))
            {
                // A origem informa uma condicao comercial, nao um codigo de fabrica.
                // Mantem vazio para nao transformar a descricao em uma chave ficticia.
                codigo = string.Empty;
                codigosFabricaAusentes++;
            }
            else if (!referenciasUsadas.Add(codigo))
            {
                // Duplicidades de codigos reais ainda recebem uma chave deterministica.
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
                Ncm = DeterminarNcm(codigoOrigem, aplicacao, nome),
                UfOrigem = _fiscalParameters.UfFabrica,
                // O preco de representacao e preservado provisoriamente nas duas colunas;
                // frete, margem e tributacao ainda nao foram homologados.
                PrecoVenda = preco,
                PrecoFabrica = preco,
                Unidade = unidade,
                // Orientação da Mônica: sem confirmação comercial da embalagem,
                // manter uma unidade de venda para todos os itens.
                QtdeEmbalagemVenda = 1,
                PesoBruto = LerDecimal(worksheet.Cell(row, 14)),
                QtdeEmbalagemCompra = 1,
                UnidFabril = unidade,
                EstoqueMinimo = 0,
                EstoqueMaximo = 0,
                Observacao = string.Empty,
                SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
            };
            produto.SituacaoCamposFiscais[13] = SituacaoCampoFiscal.Confirmado;
            AplicarTributacao(produto, tributacaoSt);
            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Nina Martinelli processados.");
        Console.WriteLine($"[INFO] {referenciasTecnicasCriadas} referencia(s) tecnica(s) criada(s) para codigos repetidos.");
        Console.WriteLine($"[INFO] {codigosFabricaAusentes} produto(s) sem codigo de fabrica na origem; campo mantido vazio.");
        Console.WriteLine("[INFO] Fiscal aplicado conforme orientação da Mônica: CSOSN 500, CST 000 e CFOP 5405/6404.");
        Console.WriteLine("[INFO] NCM aplicado por familia: ceramicos 69072300, cimenticios 68101900 e impermeabilizantes 38099390.");
        return Task.FromResult(produtos);
    }

    private List<ProdutoErpDto> LerImportacaoErp(IXLWorksheet worksheet)
    {
        var produtos = new List<ProdutoErpDto>();
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var row = 2; row <= ultimaLinha; row++)
        {
            if (string.IsNullOrWhiteSpace(worksheet.Cell(row, 2).GetFormattedString())
                && string.IsNullOrWhiteSpace(worksheet.Cell(row, 5).GetFormattedString()))
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
            produto.Ncm = DeterminarNcm(produto.CodigoFabrica, produto.Linha, produto.DescricaoComercial);
            produto.Observacao = string.Empty;
            produto.SituacaoCamposFiscais[13] = SituacaoCampoFiscal.Confirmado;
            AplicarTributacao(
                produto,
                new TributacaoStLinha(
                    produto.PercentualSt,
                    produto.PercentualSt == 0m,
                    produto.Csosn == "500",
                    produto.Csosn == "400"));
            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos Nina Martinelli atualizados a partir de uma importacao ERP existente.");
        return produtos;
    }

    private void AplicarTributacao(ProdutoErpDto produto, TributacaoStLinha tributacaoSt)
    {
        produto.UfOrigem = _fiscalParameters.UfFabrica;
        produto.IpiPercentual = _fiscalParameters.Ipi;
        produto.AliqIcmsOrigem = _fiscalParameters.AliquotaIcmsOrigem;
        produto.AliqIcmsInterna = _fiscalParameters.AliquotaIcmsSaida;
        produto.Iva = _fiscalParameters.Mva;
        produto.PercentualSt = _fiscalParameters.PercentualSt;
        produto.Cst = _fiscalParameters.Cst;
        produto.AliquotaCofinsCst = _fiscalParameters.AliquotaCofinsCst;
        produto.AliquotaIpiCst = _fiscalParameters.AliquotaIpiCst;
        produto.AliquotaPisCst = _fiscalParameters.AliquotaPisCst;
        produto.Csosn = _fiscalParameters.Csosn;
        produto.CfopDentro = _fiscalParameters.CfopDentro;
        produto.CfopFora = _fiscalParameters.CfopFora;
        produto.EnquadramentoIpi = _fiscalParameters.EnquadramentoIpi;
        produto.AliquotaPisOrigem = _fiscalParameters.PisOrigemTexto;
        produto.AliquotaCofinsOrigem = _fiscalParameters.CofinsOrigemTexto;
        produto.AliquotaIbs = _fiscalParameters.AliquotaIbs;
        produto.AliquotaCbs = _fiscalParameters.AliquotaCbs;
        produto.ClassificacaoTributaria = _fiscalParameters.ClassificacaoTributaria;

        foreach (var coluna in new[]
                 {
                     14, 18, 19, 20, 21, 26, 27, 28, 29, 30, 31, 32, 39,
                     51, 52, 53, 57, 58, 59
                 })
        {
            produto.SituacaoCamposFiscais[coluna] = SituacaoCampoFiscal.Confirmado;
        }
    }

    private static TributacaoStLinha LerTributacaoSt(IXLWorksheet worksheet, int row)
    {
        var textoLinha = string.Join(
            " ",
            worksheet.Row(row).Cells(1, worksheet.LastColumnUsed()?.ColumnNumber() ?? 22)
                .Select(cell => cell.GetFormattedString()));
        var textoNormalizado = NormalizarTexto(textoLinha);

        if (textoNormalizado.Contains("ST RECOLHID", StringComparison.Ordinal)
            || textoNormalizado.Contains("ST RETID", StringComparison.Ordinal)
            || textoNormalizado.Contains("ST COBRADO ANTERIORMENTE", StringComparison.Ordinal)
            || textoNormalizado.Contains("ST COBRADA ANTERIORMENTE", StringComparison.Ordinal))
        {
            return new TributacaoStLinha(0.0986m, false, true);
        }

        if (textoNormalizado.Contains("NAO TRIBUTAD", StringComparison.Ordinal)
            || textoNormalizado.Contains("NÃO TRIBUTAD", StringComparison.Ordinal))
        {
            return new TributacaoStLinha(0m, true, OperacaoNaoTributada: true);
        }

        if (textoNormalizado.Contains("SEM ST", StringComparison.Ordinal)
            || textoNormalizado.Contains("ST ISENTO", StringComparison.Ordinal)
            || textoNormalizado.Contains("NAO TEM ST", StringComparison.Ordinal)
            || textoNormalizado.Contains("NÃO TEM ST", StringComparison.Ordinal))
        {
            return new TributacaoStLinha(0m, true);
        }

        var match = Regex.Match(
            textoNormalizado,
            @"(?:PERCENTUAL\s*)?ST\s*[:=-]?\s*(9[,.]86|0[,.]0986)\s*%?",
            RegexOptions.IgnoreCase);
        if (match.Success)
        {
            return new TributacaoStLinha(DecimalParser.Parse(match.Groups[1].Value), false);
        }

        // Sem indicação explícita, não presume incidência de ST.
        return new TributacaoStLinha(0m, true);
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

    private string DeterminarNcm(string codigoFabrica, string aplicacao, string nome)
    {
        var codigoNormalizado = NormalizarTexto(codigoFabrica);
        var aplicacaoNormalizada = NormalizarTexto(aplicacao);
        var nomeNormalizado = NormalizarTexto(nome);

        // Os produtos de tratamento/impermeabilizacao pertencem a familia 157
        // e prevalecem sobre a empresa emissora, conforme orientacao fiscal.
        if (codigoNormalizado.StartsWith("157", StringComparison.Ordinal)
            || aplicacaoNormalizada.Contains("QUIMICO", StringComparison.Ordinal))
        {
            return _fiscalParameters.NcmImpermeabilizantes;
        }

        // Na tabela do fornecedor, 154, 158 e 159 identificam as linhas da
        // Nina Martinelli Revestimentos (cimenticios). As referencias textuais
        // "depende do raio" sao bordas Verano inseridas na mesma familia 154.
        if (PrefixosProdutosCimenticios.Any(prefixo =>
                codigoNormalizado.StartsWith(prefixo, StringComparison.Ordinal))
            || codigoNormalizado.Equals("DEPENDE DO RAIO", StringComparison.Ordinal)
            || ((string.IsNullOrEmpty(codigoNormalizado)
                    || codigoNormalizado.StartsWith("BORDA-", StringComparison.Ordinal))
                && nomeNormalizado.Contains("BORDA", StringComparison.Ordinal)
                && nomeNormalizado.Contains("VERANO", StringComparison.Ordinal)))
        {
            return _fiscalParameters.NcmProdutosCimenticios;
        }

        return _fiscalParameters.NcmProdutosCeramicos;
    }

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

    private sealed record TributacaoStLinha(
        decimal PercentualInformado,
        bool SemIncidencia,
        bool IcmsStCobradoAnteriormente = false,
        bool OperacaoNaoTributada = false);
}
