using System.Security.Cryptography;
using System.Text.Json;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Persistence;

public sealed class AdamaAtualizacaoService(
    IAdamaReader adamaReader,
    R3IntegradorDbInitializer dbInitializer,
    IExcelResumidoExporter excelResumidoExporter)
{
    private const string TabelaOrigem = "ADAMA";
    private const int QuantidadeColunasErp = 60;

    public async Task<ResultadoAtualizacaoAdama> ExecutarAsync(
        string caminhoTabelaFornecedor,
        string caminhoPlanilhaContador,
        string caminhoBanco,
        string pastaSaida)
    {
        caminhoTabelaFornecedor = ValidarArquivo(caminhoTabelaFornecedor, "tabela do fornecedor");
        caminhoPlanilhaContador = ValidarArquivo(caminhoPlanilhaContador, "planilha validada pelo contador");
        caminhoBanco = Path.GetFullPath(caminhoBanco);
        pastaSaida = Path.GetFullPath(pastaSaida);

        if (!File.Exists(caminhoBanco))
        {
            throw new FileNotFoundException("O banco de dados informado nao foi encontrado.", caminhoBanco);
        }

        var produtosTabela = (await adamaReader.LerAsync(caminhoTabelaFornecedor))
            .ToDictionary(produto => produto.CodigoFabrica, StringComparer.OrdinalIgnoreCase);

        var identificador = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var pastaExecucao = Path.Combine(pastaSaida, $"ADAMA_PROCESSADO_{identificador}");
        Directory.CreateDirectory(pastaExecucao);

        var caminhoPlanilhaCorrigida = Path.Combine(
            pastaExecucao,
            "BASE_AUDITORIA_ADAMA_CONTADOR_OK_CORRIGIDA.xlsx");
        var alteracoes = CorrigirPlanilhaContador(
            caminhoPlanilhaContador,
            caminhoPlanilhaCorrigida,
            produtosTabela);
        var caminhoPlanilhaAtualizacao = Path.Combine(
            pastaExecucao,
            "ATUALIZACAO_ADAMA_PRECOS.xlsx");
        var referenciasAlteradas = alteracoes.Itens
            .Select(item => item.Referencia)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var produtosAtualizacao = produtosTabela.Values
            .Where(produto => referenciasAlteradas.Contains(produto.CodigoFabrica))
            .OrderBy(produto => produto.CodigoFabrica)
            .ToArray();
        await excelResumidoExporter.ExportarAsync(produtosAtualizacao, caminhoPlanilhaAtualizacao);
        if (produtosAtualizacao.Length != alteracoes.Itens.Count)
        {
            throw new InvalidDataException(
                $"A atualizacao Adama contem {produtosAtualizacao.Length} produtos; eram esperados {alteracoes.Itens.Count}.");
        }

        var caminhoBackup = Path.Combine(
            pastaExecucao,
            $"R3IntegradorDb_ANTES_ADAMA_{identificador}.db");
        File.Copy(caminhoBanco, caminhoBackup, overwrite: false);

        var importacao = dbInitializer.CriarOuAtualizar(
            caminhoBanco,
            caminhoPlanilhaCorrigida,
            TabelaOrigem);
        var produtosValidadosNoBanco = ValidarBanco(caminhoBanco, produtosTabela);

        var caminhoAuditoria = Path.Combine(pastaExecucao, "RELATORIO_AUDITORIA_ADAMA.json");
        var auditoria = new
        {
            status = "SUCESSO",
            processadoEm = DateTime.Now,
            regra = new
            {
                representacao = true,
                precoVenda = "PRECO 2026",
                precoFabrica = "PRECO 2026",
                calculo = "Nenhum acrescimo, margem ou conversao",
                valorM2 = "Ignorado para precificacao",
                importacaoInicial = "Ja realizada no ERP",
                atualizacaoErp = "Layout resumido de 8 colunas, somente para os produtos com preco divergente"
            },
            arquivos = new
            {
                tabelaFornecedor = caminhoTabelaFornecedor,
                sha256TabelaFornecedor = CalcularSha256(caminhoTabelaFornecedor),
                planilhaContador = caminhoPlanilhaContador,
                sha256PlanilhaContador = CalcularSha256(caminhoPlanilhaContador),
                planilhaCorrigida = caminhoPlanilhaCorrigida,
                sha256PlanilhaCorrigida = CalcularSha256(caminhoPlanilhaCorrigida),
                planilhaAtualizacao = caminhoPlanilhaAtualizacao,
                sha256PlanilhaAtualizacao = CalcularSha256(caminhoPlanilhaAtualizacao),
                banco = caminhoBanco,
                backupBanco = caminhoBackup
            },
            contagens = new
            {
                produtosTabela = produtosTabela.Count,
                produtosPlanilhaContador = alteracoes.ProdutosPlanilha,
                produtosComPrecoCorrigido = alteracoes.ProdutosComPrecoCorrigido,
                produtosComUnidadeCorrigida = alteracoes.ProdutosComUnidadeCorrigida,
                produtosImportadosNoBanco = importacao.ProdutosImportados,
                produtosValidadosNoBanco
            },
            alteracoes = alteracoes.Itens
        };
        await File.WriteAllTextAsync(
            caminhoAuditoria,
            JsonSerializer.Serialize(auditoria, new JsonSerializerOptions { WriteIndented = true }));

        return new ResultadoAtualizacaoAdama(
            pastaExecucao,
            caminhoPlanilhaCorrigida,
            caminhoPlanilhaAtualizacao,
            caminhoBanco,
            caminhoBackup,
            caminhoAuditoria,
            alteracoes.ProdutosPlanilha,
            alteracoes.ProdutosComPrecoCorrigido,
            alteracoes.ProdutosComUnidadeCorrigida,
            produtosValidadosNoBanco);
    }

    private static AlteracoesPlanilha CorrigirPlanilhaContador(
        string caminhoOrigem,
        string caminhoDestino,
        IReadOnlyDictionary<string, ProdutoErpDto> produtosTabela)
    {
        using var workbook = new XLWorkbook(caminhoOrigem);
        var worksheet = workbook.Worksheets.FirstOrDefault()
            ?? throw new InvalidDataException("A planilha do contador nao possui abas.");
        ValidarLayoutErp(worksheet);

        var referenciasEncontradas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var itens = new List<AlteracaoAdama>();
        var produtosPlanilha = 0;
        var produtosComPrecoCorrigido = 0;
        var produtosComUnidadeCorrigida = 0;
        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 1;

        for (var linha = 2; linha <= ultimaLinha; linha++)
        {
            if (worksheet.Row(linha).Cells(1, QuantidadeColunasErp).All(celula => celula.IsEmpty()))
            {
                continue;
            }

            produtosPlanilha++;
            var referencia = worksheet.Cell(linha, 2).GetString().Trim();
            if (string.IsNullOrWhiteSpace(referencia))
            {
                throw new InvalidDataException($"Codigo de fabrica ausente na linha {linha} da planilha do contador.");
            }

            if (!referenciasEncontradas.Add(referencia))
            {
                throw new InvalidDataException($"Codigo de fabrica duplicado na planilha do contador: {referencia}.");
            }

            if (!produtosTabela.TryGetValue(referencia, out var produtoTabela))
            {
                throw new InvalidDataException(
                    $"A referencia {referencia} da planilha do contador nao existe na tabela do fornecedor.");
            }

            var precoVendaAnterior = LerDecimal(worksheet.Cell(linha, 15));
            var precoFabricaAnterior = LerDecimal(worksheet.Cell(linha, 16));
            var unidadeAnterior = worksheet.Cell(linha, 24).GetString().Trim();
            var quantidadeAnterior = LerDecimal(worksheet.Cell(linha, 25));
            var precoCorreto = produtoTabela.PrecoVenda;
            var precoAlterado = precoVendaAnterior != precoCorreto || precoFabricaAnterior != precoCorreto;
            var unidadeAlterada = !unidadeAnterior.Equals(produtoTabela.Unidade, StringComparison.OrdinalIgnoreCase)
                || quantidadeAnterior != produtoTabela.QtdeEmbalagemVenda;

            if (precoAlterado)
            {
                produtosComPrecoCorrigido++;
            }

            if (unidadeAlterada)
            {
                produtosComUnidadeCorrigida++;
            }

            if (precoAlterado || unidadeAlterada)
            {
                itens.Add(new AlteracaoAdama(
                    referencia,
                    linha,
                    precoVendaAnterior,
                    precoFabricaAnterior,
                    precoCorreto,
                    unidadeAnterior,
                    produtoTabela.Unidade,
                    quantidadeAnterior,
                    produtoTabela.QtdeEmbalagemVenda));
            }

            worksheet.Cell(linha, 15).Value = precoCorreto;
            worksheet.Cell(linha, 16).Value = precoCorreto;
            worksheet.Cell(linha, 24).Value = produtoTabela.Unidade;
            worksheet.Cell(linha, 25).Value = produtoTabela.QtdeEmbalagemVenda;
        }

        var ausentes = produtosTabela.Keys
            .Where(referencia => !referenciasEncontradas.Contains(referencia))
            .OrderBy(referencia => referencia)
            .ToArray();
        if (ausentes.Length > 0)
        {
            throw new InvalidDataException(
                $"A planilha do contador nao contem {ausentes.Length} referencia(s) da tabela: {string.Join(", ", ausentes)}");
        }

        workbook.SaveAs(caminhoDestino);
        return new AlteracoesPlanilha(
            produtosPlanilha,
            produtosComPrecoCorrigido,
            produtosComUnidadeCorrigida,
            itens);
    }

    private static void ValidarLayoutErp(IXLWorksheet worksheet)
    {
        var ultimaColuna = worksheet.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;
        if (ultimaColuna != QuantidadeColunasErp
            || !worksheet.Cell(1, 2).GetString().Contains("FÁBRICA", StringComparison.OrdinalIgnoreCase)
            || !worksheet.Cell(1, 15).GetString().Contains("VENDA", StringComparison.OrdinalIgnoreCase)
            || !worksheet.Cell(1, 16).GetString().Contains("FÁBRICA", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidDataException("A planilha do contador nao esta no layout de 60 colunas do ERP.");
        }
    }

    private static int ValidarBanco(
        string caminhoBanco,
        IReadOnlyDictionary<string, ProdutoErpDto> produtosTabela)
    {
        using var connection = new SqliteConnection($"Data Source={caminhoBanco};Mode=ReadOnly");
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = """
            SELECT "CodigoFabrica", "PrecoVenda", "PrecoFabrica", "Unidade", "QtdeEmbalagemVenda"
            FROM "Produtos"
            WHERE "TabelaOrigem" = $origem;
            """;
        command.Parameters.AddWithValue("$origem", TabelaOrigem);

        var encontrados = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var referencia = reader.GetString(0);
            if (!encontrados.Add(referencia) || !produtosTabela.TryGetValue(referencia, out var esperado))
            {
                throw new InvalidDataException($"Referencia inesperada ou duplicada no banco: {referencia}.");
            }

            var precoVenda = Convert.ToDecimal(reader.GetDouble(1));
            var precoFabrica = Convert.ToDecimal(reader.GetDouble(2));
            var unidade = reader.GetString(3);
            var quantidadeVenda = Convert.ToDecimal(reader.GetDouble(4));
            if (Math.Abs(precoVenda - esperado.PrecoVenda) > 0.0001m
                || Math.Abs(precoFabrica - esperado.PrecoFabrica) > 0.0001m
                || !unidade.Equals(esperado.Unidade, StringComparison.OrdinalIgnoreCase)
                || quantidadeVenda != esperado.QtdeEmbalagemVenda)
            {
                throw new InvalidDataException($"Valores divergentes apos a carga no banco para {referencia}.");
            }
        }

        if (encontrados.Count != produtosTabela.Count)
        {
            throw new InvalidDataException(
                $"O banco contem {encontrados.Count} produtos Adama; eram esperados {produtosTabela.Count}.");
        }

        return encontrados.Count;
    }

    private static decimal LerDecimal(IXLCell cell)
    {
        if (cell.IsEmpty())
        {
            return 0;
        }

        if (cell.TryGetValue<decimal>(out var valor))
        {
            return valor;
        }

        throw new InvalidDataException($"Valor numerico invalido na celula {cell.Address}: {cell.GetString()}.");
    }

    private static string ValidarArquivo(string caminho, string descricao)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(caminho);
        caminho = Path.GetFullPath(caminho);
        return File.Exists(caminho)
            ? caminho
            : throw new FileNotFoundException($"A {descricao} nao foi encontrada.", caminho);
    }

    private static string CalcularSha256(string caminho)
    {
        using var stream = File.OpenRead(caminho);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private sealed record AlteracoesPlanilha(
        int ProdutosPlanilha,
        int ProdutosComPrecoCorrigido,
        int ProdutosComUnidadeCorrigida,
        IReadOnlyList<AlteracaoAdama> Itens);
}

public sealed record AlteracaoAdama(
    string Referencia,
    int Linha,
    decimal PrecoVendaAnterior,
    decimal PrecoFabricaAnterior,
    decimal PrecoCorreto,
    string UnidadeAnterior,
    string UnidadeCorreta,
    decimal QuantidadeVendaAnterior,
    decimal QuantidadeVendaCorreta);

public sealed record ResultadoAtualizacaoAdama(
    string PastaSaida,
    string CaminhoPlanilhaCorrigida,
    string CaminhoPlanilhaAtualizacao,
    string CaminhoBanco,
    string CaminhoBackupBanco,
    string CaminhoAuditoria,
    int ProdutosPlanilha,
    int ProdutosComPrecoCorrigido,
    int ProdutosComUnidadeCorrigida,
    int ProdutosValidadosNoBanco);
