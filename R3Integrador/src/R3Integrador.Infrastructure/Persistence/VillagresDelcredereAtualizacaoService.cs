using System.Globalization;
using System.Reflection;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Mappers;

namespace R3Integrador.Infrastructure.Persistence;

/// <summary>
/// Separa a tabela Del Credere Villagres entre referências já cadastradas e
/// inclusões. O SQLite é consultado somente para leitura: esta operação não
/// altera preços nem cadastros no banco.
/// </summary>
public sealed class VillagresDelcredereAtualizacaoService(
    IDelcredereReader delcredereReader,
    IExcelResumidoExporter excelResumidoExporter,
    IExcelExporter excelExporter)
{
    private const string TabelaOrigemBanco = "VAREJO";
    private const decimal AcrescimoDelCredere = 1.0536m;

    // Confirmadas na tela de retorno do ERP em 12/08/2026. Elas devem ser
    // atualizadas mesmo enquanto a réplica SQLite ainda não contém o cadastro.
    private static readonly HashSet<string> ReferenciasConfirmadasNoErp =
    [
        "828001", "828002", "828003", "828004", "828005", "828006", "828007",
        "910033", "910034",
        "920067", "920068", "920069", "920070",
        "108082", "108083", "108084", "108085", "108086", "108087", "108088",
        "120001", "120003", "120004", "120005", "120006",
        // A resposta do ERP ao arquivo DEL5 Villa Art confirmou que todas as
        // referências abaixo já existem; a existência é do produto, não da
        // faixa de preço, portanto elas atualizam DEL5 a DEL30.
        "200029", "200030", "200035",
        "800061", "800062", "800085", "800086",
        "108058", "108059", "108060", "108061", "108071", "108072",
        "123010", "123011", "123049", "123050", "123051", "123053",
        "123083", "123084", "123085", "123086", "123087"
    ];

    public async Task<ResultadoAtualizacaoDelcredereComBanco> ExecutarAsync(
        string caminhoPlanilha,
        string caminhoBanco,
        string pastaSaida)
    {
        caminhoPlanilha = Path.GetFullPath(caminhoPlanilha);
        caminhoBanco = Path.GetFullPath(caminhoBanco);
        pastaSaida = Path.GetFullPath(pastaSaida);

        if (!File.Exists(caminhoPlanilha))
        {
            throw new FileNotFoundException("A tabela Del Credere nao foi encontrada.", caminhoPlanilha);
        }

        if (!File.Exists(caminhoBanco))
        {
            throw new FileNotFoundException("O banco SQLite nao foi encontrado.", caminhoBanco);
        }

        var produtosFonte = await delcredereReader.LerAsync(caminhoPlanilha);
        var produtosBanco = LerProdutosBanco(caminhoBanco);
        var referenciasFonte = produtosFonte
            .GroupBy(produto => produto.Referencia, StringComparer.OrdinalIgnoreCase)
            .Select(grupo => grupo.First())
            .ToList();

        var referenciasExistentesNoBanco = referenciasFonte
            .Where(produto => produtosBanco.ContainsKey(produto.Referencia))
            .Select(produto => produto.Referencia)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var referenciasConfirmadasNoErp = referenciasFonte
            .Where(produto => ReferenciasConfirmadasNoErp.Contains(produto.Referencia))
            .Select(produto => produto.Referencia)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var referenciasExistentes = referenciasExistentesNoBanco
            .Union(referenciasConfirmadasNoErp, StringComparer.OrdinalIgnoreCase)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        var referenciasNovas = referenciasFonte
            .Where(produto => !referenciasExistentes.Contains(produto.Referencia))
            .Select(produto => produto.Referencia)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var identificadorExecucao = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var pastaExecucao = Path.Combine(
            pastaSaida,
            $"ATUALIZACAO_E_INCLUSAO_DELCREDERE_VILLAGRES_{identificadorExecucao}");
        Directory.CreateDirectory(pastaExecucao);

        var arquivosAtualizacao = new List<string>();
        var arquivosInclusao = new List<string>();
        var modelosFiscaisUsados = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var marcas = produtosFonte
            .Select(produto => produto.Marca)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(marca => marca, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var grupoTabela in produtosFonte
            .GroupBy(produto => produto.TabelaPreco)
            .OrderBy(grupo => grupo.Key))
        {
            foreach (var marca in marcas)
            {
                var produtosMarca = grupoTabela
                    .Where(produto => produto.Marca.Equals(marca, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                var existentes = produtosMarca
                    .Where(produto => referenciasExistentes.Contains(produto.Referencia))
                    .Select(produto => produtosBanco.TryGetValue(produto.Referencia, out var produtoBanco)
                        ? CriarProdutoAtualizacao(produto, produtoBanco)
                        : CriarProdutoAtualizacaoConfirmadaNoErp(produto))
                    .ToList();
                var novos = produtosMarca
                    .Where(produto => referenciasNovas.Contains(produto.Referencia))
                    .Select(produto => CriarProdutoInclusao(produto, ObterModeloFiscal(produto, produtosBanco, modelosFiscaisUsados)))
                    .ToList();

                ValidarPrecos(grupoTabela.Key, existentes, novos);

                var marcaArquivo = SanitizarMarcaArquivo(marca);
                if (existentes.Count > 0)
                {
                    var caminhoAtualizacao = Path.Combine(
                        pastaExecucao,
                        $"ATUALIZACAO_DELCREDERE_{grupoTabela.Key}_{marcaArquivo}_CADASTRADOS.xlsx");
                    await excelResumidoExporter.ExportarAsync(existentes, caminhoAtualizacao);
                    arquivosAtualizacao.Add(caminhoAtualizacao);
                }

                if (novos.Count > 0)
                {
                    var caminhoInclusao = Path.Combine(
                        pastaExecucao,
                        $"IMPORTACAO_ERP_DELCREDERE_{grupoTabela.Key}_{marcaArquivo}_NOVOS.xlsx");
                    await excelExporter.ExportarAsync(novos, caminhoInclusao);
                    arquivosInclusao.Add(caminhoInclusao);
                }
            }
        }

        var referenciasBancoForaDaTabela = produtosBanco.Keys
            .Where(referencia => !referenciasExistentes.Contains(referencia)
                && !referenciasNovas.Contains(referencia))
            .OrderBy(referencia => referencia, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var caminhoAuditoria = Path.Combine(pastaExecucao, "RELATORIO_AUDITORIA_ATUALIZACAO_E_INCLUSAO.json");
        var auditoria = new
        {
            status = "SUCESSO",
            processadoEm = DateTime.Now,
            arquivoOrigem = caminhoPlanilha,
            bancoConsultado = caminhoBanco,
            bancoAlterado = false,
            referenciasFonte = referenciasFonte.Count,
            referenciasParaAtualizacao = referenciasExistentes.Count,
            referenciasParaAtualizacaoConfirmadasNoErp = referenciasConfirmadasNoErp.Count,
            referenciasParaAtualizacaoEncontradasNoBanco = referenciasExistentesNoBanco.Count,
            referenciasParaInclusao = referenciasNovas.Count,
            referenciasNoBancoAusentesNaTabela = referenciasBancoForaDaTabela.Length,
            listaReferenciasParaInclusao = referenciasNovas.OrderBy(referencia => referencia).ToArray(),
            listaReferenciasConfirmadasNoErp = referenciasConfirmadasNoErp.OrderBy(referencia => referencia).ToArray(),
            listaReferenciasNoBancoAusentesNaTabela = referenciasBancoForaDaTabela,
            modeloFiscalPorGrupoNcm = modelosFiscaisUsados,
            porTabela = produtosFonte
                .GroupBy(produto => produto.TabelaPreco)
                .OrderBy(grupo => grupo.Key)
                .ToDictionary(
                    grupo => grupo.Key,
                    grupo => new
                    {
                        atualizacoes = grupo.Count(produto => referenciasExistentes.Contains(produto.Referencia)),
                        inclusoes = grupo.Count(produto => referenciasNovas.Contains(produto.Referencia)),
                        precosInvalidos = grupo.Count(produto => produto.PrecoTabela <= 0m),
                        calculosVendaInvalidos = grupo.Count(produto =>
                            Math.Round(produto.PrecoTabela * AcrescimoDelCredere, 2) != produto.PrecoVenda)
                    }),
            porMarcaETabela = produtosFonte
                .GroupBy(produto => new { produto.TabelaPreco, produto.Marca })
                .OrderBy(grupo => grupo.Key.TabelaPreco)
                .ThenBy(grupo => grupo.Key.Marca)
                .ToDictionary(
                    grupo => $"{grupo.Key.TabelaPreco} - {grupo.Key.Marca}",
                    grupo => new
                    {
                        atualizacoes = grupo.Count(produto => referenciasExistentes.Contains(produto.Referencia)),
                        inclusoes = grupo.Count(produto => referenciasNovas.Contains(produto.Referencia))
                    })
        };
        await File.WriteAllTextAsync(
            caminhoAuditoria,
            JsonSerializer.Serialize(auditoria, new JsonSerializerOptions { WriteIndented = true }));

        return new ResultadoAtualizacaoDelcredereComBanco(
            pastaExecucao,
            arquivosAtualizacao,
            arquivosInclusao,
            caminhoAuditoria,
            referenciasExistentes.Count,
            referenciasNovas.Count,
            referenciasBancoForaDaTabela.Length);
    }

    /// <summary>
    /// Inclui na réplica local somente referências cuja existência já foi
    /// confirmada no ERP, sem sobrescrever qualquer cadastro existente.
    /// </summary>
    public async Task<ResultadoSincronizacaoErp> SincronizarConfirmadosNoErpAsync(
        string caminhoPlanilha,
        string caminhoBanco,
        string pastaSaida)
    {
        caminhoPlanilha = Path.GetFullPath(caminhoPlanilha);
        caminhoBanco = Path.GetFullPath(caminhoBanco);
        pastaSaida = Path.GetFullPath(pastaSaida);
        Directory.CreateDirectory(pastaSaida);

        if (!File.Exists(caminhoPlanilha) || !File.Exists(caminhoBanco))
        {
            throw new FileNotFoundException("A tabela Del Credere ou o banco SQLite nao foi encontrado.");
        }

        var produtosFonte = await delcredereReader.LerAsync(caminhoPlanilha);
        var produtosBanco = LerProdutosBanco(caminhoBanco);
        var produtosParaSincronizar = produtosFonte
            .Where(produto => produto.TabelaPreco.Equals("DEL5", StringComparison.OrdinalIgnoreCase))
            .Where(produto => ReferenciasConfirmadasNoErp.Contains(produto.Referencia))
            .Where(produto => !produtosBanco.ContainsKey(produto.Referencia))
            .OrderBy(produto => produto.Referencia, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var modelosFiscais = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var produtosErp = produtosParaSincronizar
            .Select(produto => CriarProdutoParaReplicaLocal(produto, ObterModeloFiscal(produto, produtosBanco, modelosFiscais)))
            .ToList();

        var backup = Path.Combine(
            pastaSaida,
            $"R3IntegradorDb_ANTES_SINCRONIZACAO_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.db.bak");
        File.Copy(caminhoBanco, backup, overwrite: false);
        InserirProdutosNoBanco(caminhoBanco, produtosErp);

        return new ResultadoSincronizacaoErp(
            backup,
            produtosErp.Count,
            produtosErp.Select(produto => produto.CodigoFabrica).ToArray(),
            modelosFiscais);
    }

    private static ProdutoErpDto CriarProdutoAtualizacao(
        ProdutoNormalizado produtoFonte,
        ProdutoErpDto produtoBanco)
    {
        var produto = CopiarProduto(produtoBanco);
        produto.Marca = CriarMarcaTabela(produtoFonte.Marca, produtoFonte.TabelaPreco);
        produto.PrecoFabrica = produtoFonte.PrecoTabela;
        produto.PrecoVenda = produtoFonte.PrecoVenda;
        produto.QtdeEmbalagemVenda = produto.Unidade.Equals("M2", StringComparison.OrdinalIgnoreCase) ? 0m : produto.QtdeEmbalagemVenda;
        AdicionarSegmentacaoDescricao(produto, produtoFonte.SegmentacaoComercial);
        return produto;
    }

    private static ProdutoErpDto CriarProdutoAtualizacaoConfirmadaNoErp(ProdutoNormalizado produtoFonte)
    {
        // O arquivo reduzido precisa somente dos dados comerciais. O cadastro
        // completo está no ERP e não deve ser recriado/substituído pelo SQLite.
        var produto = ProdutoErpMapper.Map(produtoFonte);
        produto.Marca = CriarMarcaTabela(produtoFonte.Marca, produtoFonte.TabelaPreco);
        produto.PrecoFabrica = produtoFonte.PrecoTabela;
        produto.PrecoVenda = produtoFonte.PrecoVenda;
        produto.QtdeEmbalagemVenda = produto.Unidade.Equals("M2", StringComparison.OrdinalIgnoreCase) ? 0m : produto.QtdeEmbalagemVenda;
        AdicionarSegmentacaoDescricao(produto, produtoFonte.SegmentacaoComercial);
        return produto;
    }

    private static ProdutoErpDto CriarProdutoInclusao(
        ProdutoNormalizado produtoFonte,
        ProdutoErpDto modeloFiscal)
    {
        var produto = ProdutoErpMapper.Map(produtoFonte);
        produto.Marca = CriarMarcaTabela(produtoFonte.Marca, produtoFonte.TabelaPreco);
        CopiarCamposFiscais(modeloFiscal, produto);
        produto.PrecoFabrica = produtoFonte.PrecoTabela;
        produto.PrecoVenda = produtoFonte.PrecoVenda;
        produto.QtdeEmbalagemVenda = produto.Unidade.Equals("M2", StringComparison.OrdinalIgnoreCase) ? 0m : produto.QtdeEmbalagemVenda;
        AdicionarSegmentacaoDescricao(produto, produtoFonte.SegmentacaoComercial);
        return produto;
    }

    private static ProdutoErpDto CriarProdutoParaReplicaLocal(
        ProdutoNormalizado produtoFonte,
        ProdutoErpDto modeloFiscal)
    {
        var produto = ProdutoErpMapper.Map(produtoFonte);
        // A réplica VAREJO guarda a marca-base; o sufixo da faixa é aplicado
        // somente ao exportar uma atualização Del Credere.
        produto.Marca = produtoFonte.Marca;
        CopiarCamposFiscais(modeloFiscal, produto);
        produto.PrecoFabrica = produtoFonte.PrecoTabela;
        produto.PrecoVenda = produtoFonte.PrecoVenda;
        produto.QtdeEmbalagemVenda = produto.Unidade.Equals("M2", StringComparison.OrdinalIgnoreCase) ? 0m : produto.QtdeEmbalagemVenda;
        AdicionarSegmentacaoDescricao(produto, produtoFonte.SegmentacaoComercial);
        return produto;
    }

    private static void AdicionarSegmentacaoDescricao(ProdutoErpDto produto, string segmentacao)
    {
        if (!string.IsNullOrWhiteSpace(segmentacao)
            && !produto.DescricaoCompleta.Contains("SEGMENTO:", StringComparison.OrdinalIgnoreCase))
        {
            produto.DescricaoCompleta = $"{produto.DescricaoCompleta} - SEGMENTO: {segmentacao}";
        }
    }

    private static string CriarMarcaTabela(string marca, string tabelaPreco)
    {
        var faixa = tabelaPreco.Replace("DEL", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
        return string.IsNullOrWhiteSpace(faixa) ? marca : $"{marca} {faixa}";
    }

    private static string SanitizarMarcaArquivo(string marca) =>
        marca.Replace(" ", "_", StringComparison.Ordinal).ToUpperInvariant();

    private static ProdutoErpDto ObterModeloFiscal(
        ProdutoNormalizado produtoFonte,
        IReadOnlyDictionary<string, ProdutoErpDto> produtosBanco,
        IDictionary<string, string> modelosFiscaisUsados)
    {
        var grupo = produtoFonte.Grupo.Trim();
        var candidatosGrupo = produtosBanco.Values
            .Where(produto => produto.Grupo.Equals(grupo, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (candidatosGrupo.Count == 0)
        {
            throw new InvalidDataException(
                $"Nao existe produto cadastrado do grupo '{produtoFonte.Grupo}' para herdar os dados fiscais.");
        }

        var ncm = candidatosGrupo
            .Where(produto => !string.IsNullOrWhiteSpace(produto.Ncm))
            .GroupBy(produto => produto.Ncm)
            .OrderByDescending(grupoNcm => grupoNcm.Count())
            .ThenBy(grupoNcm => grupoNcm.Key, StringComparer.OrdinalIgnoreCase)
            .First().Key;
        var modelo = candidatosGrupo.First(produto => produto.Ncm.Equals(ncm, StringComparison.OrdinalIgnoreCase));
        modelosFiscaisUsados[grupo] = ncm;
        return modelo;
    }

    private static void CopiarCamposFiscais(ProdutoErpDto origem, ProdutoErpDto destino)
    {
        destino.Ncm = origem.Ncm;
        destino.UfOrigem = origem.UfOrigem;
        destino.IpiPercentual = origem.IpiPercentual;
        destino.AliqIcmsOrigem = origem.AliqIcmsOrigem;
        destino.AliqIcmsInterna = origem.AliqIcmsInterna;
        destino.Iva = origem.Iva;
        destino.Cst = origem.Cst;
        destino.AliquotaCofinsCst = origem.AliquotaCofinsCst;
        destino.AliquotaIpiCst = origem.AliquotaIpiCst;
        destino.AliquotaPisCst = origem.AliquotaPisCst;
        destino.Csosn = origem.Csosn;
        destino.CfopDentro = origem.CfopDentro;
        destino.CfopFora = origem.CfopFora;
        destino.ValorPi = origem.ValorPi;
        destino.AliquotaCofins = origem.AliquotaCofins;
        destino.AliquotaPis = origem.AliquotaPis;
        destino.PercentualSt = origem.PercentualSt;
        destino.DiferencaIcms = origem.DiferencaIcms;
        destino.ReducaoBaseIcms = origem.ReducaoBaseIcms;
        destino.ReducaoBaseSt = origem.ReducaoBaseSt;
        destino.EnquadramentoIpi = origem.EnquadramentoIpi;
        destino.AliquotaPisOrigem = origem.AliquotaPisOrigem;
        destino.AliquotaCofinsOrigem = origem.AliquotaCofinsOrigem;
        destino.AliquotaIbs = origem.AliquotaIbs;
        destino.AliquotaCbs = origem.AliquotaCbs;
        destino.ClassificacaoTributaria = origem.ClassificacaoTributaria;
        destino.CodigoBeneficio = origem.CodigoBeneficio;

        foreach (var coluna in new[] { 13, 14, 18, 19, 20, 21, 26, 27, 28, 29, 30, 31, 32, 36, 37, 38, 39, 42, 43, 44, 51, 52, 53, 57, 58, 59, 60 })
        {
            destino.SituacaoCamposFiscais[coluna] = SituacaoCampoFiscal.Confirmado;
        }
    }

    private static void ValidarPrecos(string tabelaPreco, IEnumerable<ProdutoErpDto> existentes, IEnumerable<ProdutoErpDto> novos)
    {
        foreach (var produto in existentes.Concat(novos))
        {
            if (produto.PrecoFabrica <= 0m)
            {
                throw new InvalidDataException($"Preco de fabrica invalido em {tabelaPreco}: {produto.CodigoFabrica}.");
            }

            if (Math.Round(produto.PrecoFabrica * AcrescimoDelCredere, 2) != produto.PrecoVenda)
            {
                throw new InvalidDataException($"Calculo Del Credere invalido em {tabelaPreco}: {produto.CodigoFabrica}.");
            }
        }
    }

    private static Dictionary<string, ProdutoErpDto> LerProdutosBanco(string caminhoBanco)
    {
        using var connection = new SqliteConnection($"Data Source={caminhoBanco};Mode=ReadOnly");
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM \"Produtos\" WHERE \"TabelaOrigem\" = $origem;";
        command.Parameters.AddWithValue("$origem", TabelaOrigemBanco);
        using var reader = command.ExecuteReader();

        var propriedades = typeof(ProdutoErpDto).GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(propriedade => propriedade.CanWrite && propriedade.Name != nameof(ProdutoErpDto.SituacaoCamposFiscais))
            .ToArray();
        var produtos = new Dictionary<string, ProdutoErpDto>(StringComparer.OrdinalIgnoreCase);

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
                    propriedade.SetValue(produto, Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
                }
                else if (propriedade.PropertyType == typeof(string))
                {
                    propriedade.SetValue(produto, Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty);
                }
            }

            if (!produtos.TryAdd(produto.CodigoFabrica, produto))
            {
                throw new InvalidDataException($"Codigo de fabrica duplicado no banco: {produto.CodigoFabrica}.");
            }
        }

        return produtos;
    }

    private static void InserirProdutosNoBanco(string caminhoBanco, IReadOnlyCollection<ProdutoErpDto> produtos)
    {
        using var connection = new SqliteConnection($"Data Source={caminhoBanco}");
        connection.Open();
        using var transaction = connection.BeginTransaction();
        var propriedades = typeof(ProdutoErpDto).GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(propriedade => propriedade.CanRead
                && propriedade.Name != nameof(ProdutoErpDto.SituacaoCamposFiscais)
                && (propriedade.PropertyType == typeof(string) || propriedade.PropertyType == typeof(decimal)))
            .ToArray();
        var colunas = string.Join(", ", propriedades.Select(propriedade => $"\"{propriedade.Name}\""));
        var parametros = string.Join(", ", propriedades.Select((_, indice) => $"$p{indice}"));
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = $"""
            INSERT INTO "Produtos" ({colunas}, "TabelaOrigem")
            VALUES ({parametros}, $tabelaOrigem);
            """;
        foreach (var (propriedade, indice) in propriedades.Select((propriedade, indice) => (propriedade, indice)))
        {
            command.Parameters.Add(new SqliteParameter($"$p{indice}", DBNull.Value));
        }

        command.Parameters.AddWithValue("$tabelaOrigem", TabelaOrigemBanco);
        foreach (var produto in produtos)
        {
            for (var indice = 0; indice < propriedades.Length; indice++)
            {
                command.Parameters[indice].Value = propriedades[indice].GetValue(produto) ?? DBNull.Value;
            }

            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    private static ProdutoErpDto CopiarProduto(ProdutoErpDto origem)
    {
        var copia = new ProdutoErpDto();
        foreach (var propriedade in typeof(ProdutoErpDto).GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(propriedade => propriedade.CanWrite && propriedade.Name != nameof(ProdutoErpDto.SituacaoCamposFiscais)))
        {
            propriedade.SetValue(copia, propriedade.GetValue(origem));
        }

        copia.SituacaoCamposFiscais = new Dictionary<int, SituacaoCampoFiscal>(origem.SituacaoCamposFiscais);
        return copia;
    }
}

public sealed record ResultadoAtualizacaoDelcredereComBanco(
    string PastaSaida,
    IReadOnlyCollection<string> ArquivosAtualizacao,
    IReadOnlyCollection<string> ArquivosInclusao,
    string CaminhoAuditoria,
    int ReferenciasAtualizacao,
    int ReferenciasInclusao,
    int ReferenciasBancoAusentesNaTabela);

public sealed record ResultadoSincronizacaoErp(
    string CaminhoBackup,
    int ProdutosInseridos,
    IReadOnlyCollection<string> ReferenciasInseridas,
    IReadOnlyDictionary<string, string> ModeloFiscalPorGrupoNcm);
