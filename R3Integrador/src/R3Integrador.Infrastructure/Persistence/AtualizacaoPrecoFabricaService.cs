using System.Globalization;
using System.Reflection;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Mappers;

namespace R3Integrador.Infrastructure.Persistence;

public sealed class AtualizacaoPrecoFabricaService(
    IExcelReader excelReader,
    IExcelExporter excelExporter)
{
    private const string TabelaOrigem = "VAREJO";

    private static readonly int[] ColunasFiscaisPendentes =
    [
        13, 18, 19, 20, 21, 26, 27, 28, 29, 30, 31, 32, 36, 37, 38, 39,
        42, 43, 44, 45, 46, 47, 48, 49, 51, 52, 53, 57, 58, 59, 60
    ];

    public async Task<ResultadoAtualizacaoPrecoFabrica> ExecutarAsync(
        string caminhoBanco,
        string caminhoPlanilha,
        string pastaSaida)
    {
        caminhoBanco = Path.GetFullPath(caminhoBanco);
        caminhoPlanilha = Path.GetFullPath(caminhoPlanilha);
        pastaSaida = Path.GetFullPath(pastaSaida);

        if (!File.Exists(caminhoBanco))
        {
            throw new FileNotFoundException("O banco SQLite nao foi encontrado.", caminhoBanco);
        }

        if (!File.Exists(caminhoPlanilha))
        {
            throw new FileNotFoundException("A planilha de precos nao foi encontrada.", caminhoPlanilha);
        }

        Directory.CreateDirectory(pastaSaida);

        var precosPlanilha = LerPrecosTabela(caminhoPlanilha);
        var produtosPlanilha = await excelReader.LerAsync(caminhoPlanilha, TabelaOrigem);
        var normalizadosPorReferencia = produtosPlanilha
            .GroupBy(produto => produto.Referencia.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToDictionary(grupo => grupo.Key, grupo => grupo.First(), StringComparer.OrdinalIgnoreCase);

        var produtosBanco = LerProdutosBanco(caminhoBanco);
        var produtosAtualizacao = new List<ProdutoErpDto>();
        var produtosInclusao = new List<ProdutoErpDto>();
        var precosParaAtualizar = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        foreach (var (referencia, precoFabrica) in precosPlanilha)
        {
            if (produtosBanco.TryGetValue(referencia, out var produtoExistente))
            {
                produtoExistente.PrecoFabrica = precoFabrica;
                if (produtoExistente.Unidade.Equals("M2", StringComparison.OrdinalIgnoreCase))
                {
                    produtoExistente.QtdeEmbalagemVenda = 0;
                }

                produtosAtualizacao.Add(produtoExistente);
                precosParaAtualizar[referencia] = precoFabrica;
                continue;
            }

            if (!normalizadosPorReferencia.TryGetValue(referencia, out var produtoNormalizado))
            {
                throw new InvalidDataException(
                    $"A referencia {referencia} possui preco, mas nao pode ser lida como produto.");
            }

            var produtoInclusao = ProdutoErpMapper.Map(produtoNormalizado);
            PrepararParaInclusao(produtoInclusao, precoFabrica);
            produtosInclusao.Add(produtoInclusao);
        }

        var agora = DateTime.Now;
        var caminhoAtualizacao = Path.Combine(
            pastaSaida,
            $"IMPORTACAO_ERP_ATUALIZACAO_PRECO_FABRICA_VAREJO_{agora:yyyyMMdd_HHmmss}.xlsx");
        var caminhoInclusao = Path.Combine(
            pastaSaida,
            $"IMPORTACAO_ERP_INCLUSAO_VAREJO_PENDENTE_FISCAL_{agora:yyyyMMdd_HHmmss}.xlsx");

        await excelExporter.ExportarAsync(produtosAtualizacao, caminhoAtualizacao);
        await excelExporter.ExportarAsync(produtosInclusao, caminhoInclusao);

        var quantidadeAtualizada = AtualizarPrecosBanco(caminhoBanco, precosParaAtualizar);
        if (quantidadeAtualizada != produtosAtualizacao.Count)
        {
            throw new InvalidOperationException(
                $"Foram exportados {produtosAtualizacao.Count} produtos, mas apenas " +
                $"{quantidadeAtualizada} precos foram atualizados no banco.");
        }

        return new ResultadoAtualizacaoPrecoFabrica(
            caminhoAtualizacao,
            caminhoInclusao,
            produtosAtualizacao.Count,
            produtosInclusao.Count,
            quantidadeAtualizada,
            produtosInclusao.Select(produto => produto.CodigoFabrica).ToArray());
    }

    private static Dictionary<string, decimal> LerPrecosTabela(string caminhoPlanilha)
    {
        using var workbook = new XLWorkbook(caminhoPlanilha);
        var worksheet = workbook.Worksheet(TabelaOrigem);
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        var precos = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);

        for (var linha = 3; linha <= ultimaLinha; linha++)
        {
            var referencia = worksheet.Cell(linha, 2).GetString().Trim();
            if (string.IsNullOrWhiteSpace(referencia) || !referencia.All(char.IsDigit))
            {
                continue;
            }

            if (!worksheet.Cell(linha, 18).TryGetValue<decimal>(out var precoTabela))
            {
                throw new InvalidDataException(
                    $"Preco de tabela invalido para a referencia {referencia}, linha {linha}.");
            }

            if (!precos.TryAdd(
                referencia,
                Math.Round(precoTabela, 2, MidpointRounding.AwayFromZero)))
            {
                throw new InvalidDataException($"Referencia duplicada na planilha: {referencia}.");
            }
        }

        return precos;
    }

    private static Dictionary<string, ProdutoErpDto> LerProdutosBanco(string caminhoBanco)
    {
        using var connection = new SqliteConnection($"Data Source={caminhoBanco};Mode=ReadOnly");
        connection.Open();

        using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM \"Produtos\" WHERE \"TabelaOrigem\" = $origem;";
        command.Parameters.AddWithValue("$origem", TabelaOrigem);

        using var reader = command.ExecuteReader();
        var produtos = new Dictionary<string, ProdutoErpDto>(StringComparer.OrdinalIgnoreCase);
        var propriedades = typeof(ProdutoErpDto).GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(propriedade => propriedade.CanWrite && propriedade.Name != nameof(ProdutoErpDto.SituacaoCamposFiscais))
            .ToArray();

        while (reader.Read())
        {
            var produto = new ProdutoErpDto();
            foreach (var propriedade in propriedades)
            {
                var ordinal = reader.GetOrdinal(propriedade.Name);
                if (reader.IsDBNull(ordinal))
                {
                    continue;
                }

                if (propriedade.PropertyType == typeof(decimal))
                {
                    propriedade.SetValue(
                        produto,
                        Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
                }
                else if (propriedade.PropertyType == typeof(string))
                {
                    propriedade.SetValue(
                        produto,
                        Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty);
                }
            }

            if (!produtos.TryAdd(produto.CodigoFabrica, produto))
            {
                throw new InvalidDataException(
                    $"Codigo de fabrica duplicado no banco: {produto.CodigoFabrica}.");
            }
        }

        return produtos;
    }

    private static int AtualizarPrecosBanco(
        string caminhoBanco,
        IReadOnlyDictionary<string, decimal> precos)
    {
        using var connection = new SqliteConnection($"Data Source={caminhoBanco}");
        connection.Open();
        using var transaction = connection.BeginTransaction();
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            UPDATE "Produtos"
            SET "PrecoFabrica" = $precoFabrica
            WHERE "TabelaOrigem" = $origem AND "CodigoFabrica" = $codigoFabrica;
            """;
        var parametroPreco = command.Parameters.Add("$precoFabrica", SqliteType.Real);
        var parametroOrigem = command.Parameters.AddWithValue("$origem", TabelaOrigem);
        var parametroCodigo = command.Parameters.Add("$codigoFabrica", SqliteType.Text);
        var quantidade = 0;

        foreach (var (codigoFabrica, precoFabrica) in precos)
        {
            parametroPreco.Value = Convert.ToDouble(precoFabrica, CultureInfo.InvariantCulture);
            parametroOrigem.Value = TabelaOrigem;
            parametroCodigo.Value = codigoFabrica;
            quantidade += command.ExecuteNonQuery();
        }

        transaction.Commit();
        return quantidade;
    }

    private static void PrepararParaInclusao(ProdutoErpDto produto, decimal precoFabrica)
    {
        produto.PrecoFabrica = precoFabrica;
        produto.PrecoVenda = 0;
        produto.DescontoPercentual = 0;
        produto.QtdeEmbalagemVenda = 0;
        produto.Ncm = string.Empty;
        produto.IpiPercentual = 0;
        produto.AliqIcmsOrigem = 0;
        produto.AliqIcmsInterna = 0;
        produto.Iva = 0;
        produto.Cst = string.Empty;
        produto.AliquotaCofinsCst = string.Empty;
        produto.AliquotaIpiCst = string.Empty;
        produto.AliquotaPisCst = string.Empty;
        produto.Csosn = string.Empty;
        produto.CfopDentro = string.Empty;
        produto.CfopFora = string.Empty;
        produto.ValorPi = 0;
        produto.AliquotaCofins = 0;
        produto.AliquotaPis = 0;
        produto.PercentualSt = 0;
        produto.DiferencaIcms = 0;
        produto.ReducaoBaseIcms = 0;
        produto.ReducaoBaseSt = 0;
        produto.RetencaoPis = string.Empty;
        produto.RetencaoCofins = string.Empty;
        produto.RetencaoCsll = string.Empty;
        produto.RetencaoIrrf = string.Empty;
        produto.RetencaoPrevSocial = string.Empty;
        produto.EnquadramentoIpi = string.Empty;
        produto.AliquotaPisOrigem = string.Empty;
        produto.AliquotaCofinsOrigem = string.Empty;
        produto.AliquotaIbs = string.Empty;
        produto.AliquotaCbs = string.Empty;
        produto.ClassificacaoTributaria = string.Empty;
        produto.CodigoBeneficio = string.Empty;
        produto.Observacao = "CADASTRO NOVO - PENDENTE COMPLEMENTACAO DE NCM E TRIBUTOS";
        produto.SituacaoCamposFiscais = ColunasFiscaisPendentes.ToDictionary(
            coluna => coluna,
            _ => SituacaoCampoFiscal.Pendente);
    }
}

public sealed record ResultadoAtualizacaoPrecoFabrica(
    string CaminhoPlanilhaAtualizacao,
    string CaminhoPlanilhaInclusao,
    int ProdutosExportadosParaAtualizacao,
    int ProdutosExportadosParaInclusao,
    int ProdutosAtualizadosNoBanco,
    IReadOnlyList<string> ReferenciasParaInclusao);
