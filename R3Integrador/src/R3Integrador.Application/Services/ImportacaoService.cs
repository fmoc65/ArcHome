using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Mappers;

namespace R3Integrador.Application.Services;

public class ImportacaoService
{
    private readonly IExcelReader _excelReader;
    private readonly IVinilicoReader _vinilicoReader;
    private readonly IDelcredereReader _delcredereReader;
    private readonly IVillaArtReader _villaArtReader;
    private readonly ILastraReader _lastraReader;
    private readonly IRubinettosReader _rubinettosReader;
    private readonly IExcelExporter _excelExporter;
    private readonly ILogger<ImportacaoService> _logger;
    private readonly string _pastaSaida;

    public ImportacaoService(
        ImportacaoReaderSet readers,
        IExcelExporter excelExporter,
        ILogger<ImportacaoService> logger,
        IConfiguration configuration)
    {
        _excelReader = readers.ExcelReader;
        _vinilicoReader = readers.VinilicoReader;
        _delcredereReader = readers.DelcredereReader;
        _villaArtReader = readers.VillaArtReader;
        _lastraReader = readers.LastraReader;
        _rubinettosReader = readers.RubinettosReader;
        _excelExporter = excelExporter;
        _logger = logger;
        _pastaSaida = configuration["Diretorios:PastaSaida"]
            ?? Path.Combine(AppContext.BaseDirectory, "Saida");
    }

    public async Task ProcessarPadraoAsync(string caminhoArquivo, string tabela)
    {
        _logger.LogInformation("Iniciando importacao padrao. Arquivo={Arquivo}; Aba={Aba}", caminhoArquivo, tabela);

        var produtosBrutos = await _excelReader.LerAsync(caminhoArquivo, tabela);

        if (!PossuiProdutos(produtosBrutos, tabela))
        {
            return;
        }

        RegistrarNormalizados(tabela, caminhoArquivo, produtosBrutos);

        var produtosErp = produtosBrutos.Select(ProdutoErpMapper.Map).ToList();
        var arquivoSaida = CriarCaminhoSaida($"{tabela}_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        RegistrarExportacao(tabela, produtosErp.Count, arquivoSaida);
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Exportacao da tabela {Tabela} concluida com sucesso.", tabela);
    }

    public async Task ProcessarDelcredereAsync(string caminhoArquivo)
    {
        const string tabela = "COM DEL CREDERE";
        _logger.LogInformation("Iniciando importacao Del Credere. Arquivo={Arquivo}; Aba={Aba}", caminhoArquivo, tabela);

        var produtosBrutos = await _delcredereReader.LerAsync(caminhoArquivo);

        if (!PossuiProdutos(produtosBrutos, tabela))
        {
            return;
        }

        RegistrarNormalizados(tabela, caminhoArquivo, produtosBrutos);

        foreach (var grupoTabelaMarca in produtosBrutos
            .GroupBy(p => new { p.TabelaPreco, p.Marca })
            .OrderBy(g => g.Key.TabelaPreco)
            .ThenBy(g => g.Key.Marca))
        {
            var produtosErp = grupoTabelaMarca.Select(ProdutoErpMapper.Map).ToList();
            var sufixoTabela = string.IsNullOrWhiteSpace(grupoTabelaMarca.Key.TabelaPreco)
                ? "SEM_TABELA"
                : grupoTabelaMarca.Key.TabelaPreco;
            var sufixoMarca = SanitizarNomeArquivo(grupoTabelaMarca.Key.Marca);
            var arquivoSaida = CriarCaminhoSaida($"IMPORTACAO_ERP_DELCREDERE_{sufixoTabela}_{sufixoMarca}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            foreach (var produto in produtosErp)
            {
                produto.Voltagem = string.Empty;
            }

            RegistrarExportacao($"{tabela} {sufixoTabela} {grupoTabelaMarca.Key.Marca}", produtosErp.Count, arquivoSaida);
            await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);
        }

        _logger.LogInformation("Processamento e exportacao Del Credere concluidos com sucesso.");
    }

    public async Task ProcessarVillaArtAsync(string caminhoArquivo)
    {
        const string tabela = "VILLA ART - BOUTIQUE";
        _logger.LogInformation("Iniciando importacao Villa Art. Arquivo={Arquivo}; Aba={Aba}", caminhoArquivo, tabela);

        var produtosBrutos = await _villaArtReader.LerAsync(caminhoArquivo);

        if (!PossuiProdutos(produtosBrutos, tabela))
        {
            return;
        }

        RegistrarNormalizados(tabela, caminhoArquivo, produtosBrutos);

        var produtosErp = produtosBrutos.Select(ProdutoErpMapper.Map).ToList();
        var arquivoSaida = CriarCaminhoSaida($"VILLA_ART_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        RegistrarExportacao(tabela, produtosErp.Count, arquivoSaida);
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Exportacao Villa Art concluida com sucesso.");
    }

    public async Task ProcessarLastraAsync(string caminhoArquivo)
    {
        const string tabela = "LASTRA";
        _logger.LogInformation("Iniciando importacao Lastra. Arquivo={Arquivo}; Aba={Aba}", caminhoArquivo, tabela);

        var produtosBrutos = await _lastraReader.LerAsync(caminhoArquivo);

        if (!PossuiProdutos(produtosBrutos, tabela))
        {
            return;
        }

        RegistrarNormalizados(tabela, caminhoArquivo, produtosBrutos);

        var produtosErp = produtosBrutos.Select(ProdutoErpMapper.Map).ToList();
        var arquivoSaida = CriarCaminhoSaida($"LASTRA_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        RegistrarExportacao(tabela, produtosErp.Count, arquivoSaida);
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Processamento e exportacao de Lastra concluido com sucesso.");
    }

    public async Task ProcessarVinilicoAsync(string caminhoArquivo)
    {
        const string tabela = "VINILICO";
        _logger.LogInformation("Iniciando importacao Vinilico. Arquivo={Arquivo}; Aba={Aba}", caminhoArquivo, tabela);

        var produtosBrutos = await _vinilicoReader.LerAsync(caminhoArquivo);

        if (!PossuiProdutos(produtosBrutos, tabela))
        {
            return;
        }

        RegistrarNormalizados(tabela, caminhoArquivo, produtosBrutos);

        var produtosErp = produtosBrutos.Select(ProdutoErpMapper.Map).ToList();
        var arquivoSaida = CriarCaminhoSaida($"VINILICO_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        RegistrarExportacao(tabela, produtosErp.Count, arquivoSaida);
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Processamento e exportacao de Vinilicos concluido com sucesso.");
    }

    public async Task ProcessarRubinettosAsync(string caminhoArquivo)
    {
        const string tabela = "RUBINETTOS";
        _logger.LogInformation("Iniciando importacao Rubinettos. Arquivo={Arquivo}", caminhoArquivo);

        var produtosErp = await _rubinettosReader.LerAsync(caminhoArquivo);

        if (!produtosErp.Any())
        {
            _logger.LogWarning("Nenhum produto valido encontrado na tabela {Tabela}.", tabela);
            return;
        }

        foreach (var grupoMarca in produtosErp.GroupBy(p => p.Marca).OrderBy(g => g.Key))
        {
            var produtosMarca = grupoMarca.ToList();
            var sufixoMarca = SanitizarNomeArquivo(grupoMarca.Key);
            var arquivoSaida = CriarCaminhoSaida($"RUBINETTOS_ERP_{sufixoMarca}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            RegistrarExportacao($"{tabela} {grupoMarca.Key}", produtosMarca.Count, arquivoSaida);
            await _excelExporter.ExportarAsync(produtosMarca, arquivoSaida);
        }

        _logger.LogInformation("Processamento e exportacao Rubinettos concluidos com sucesso.");
    }

    private bool PossuiProdutos(List<ProdutoNormalizado>? produtos, string tabela)
    {
        if (produtos != null && produtos.Any())
        {
            return true;
        }

        _logger.LogWarning("Nenhum produto valido encontrado na aba {Tabela}.", tabela);
        return false;
    }

    private void RegistrarNormalizados(string tabela, string caminhoArquivo, List<ProdutoNormalizado> produtos)
    {
        _logger.LogInformation(
            "Planilha importada e normalizada. Arquivo={Arquivo}; Aba={Aba}; Produtos={Quantidade}",
            caminhoArquivo,
            tabela,
            produtos.Count);

        for (var i = 0; i < produtos.Count; i++)
        {
            var produto = produtos[i];
            _logger.LogDebug(
                "Normalizado {Indice}/{Total}: Ref={Referencia}; Grupo={Grupo}; SubGrupo={SubGrupo}; Marca={Marca}; Modelo={Modelo}; Formato={Formato}; Linha={Linha}; Colecao={Colecao}; Superficie={Superficie}; Cor={Cor}; TabelaPreco={TabelaPreco}; PrecoTabela={PrecoTabela}; PrecoDesconto={PrecoDesconto}; PrecoVenda={PrecoVenda}; M2Caixa={M2Caixa}; Espessura={Espessura}",
                i + 1,
                produtos.Count,
                produto.Referencia,
                produto.Grupo,
                produto.SubGrupo,
                produto.Marca,
                produto.Modelo,
                produto.Formato,
                produto.Linha,
                produto.Colecao,
                produto.Superficie,
                produto.Cor,
                produto.TabelaPreco,
                produto.PrecoTabela,
                produto.PrecoDesconto,
                produto.PrecoVenda,
                produto.M2Caixa,
                produto.Espessura);
        }
    }

    private void RegistrarExportacao(string tabela, int quantidade, string arquivoSaida)
    {
        _logger.LogInformation(
            "Exportando para layout ERP. Tabela={Tabela}; Registros={Quantidade}; ArquivoSaida={ArquivoSaida}",
            tabela,
            quantidade,
            arquivoSaida);
    }

    private string CriarCaminhoSaida(string nomeArquivo)
    {
        Directory.CreateDirectory(_pastaSaida);
        return Path.Combine(_pastaSaida, nomeArquivo);
    }

    private static string SanitizarNomeArquivo(string valor)
    {
        var nome = string.Join("_", valor.Split(Path.GetInvalidFileNameChars(), StringSplitOptions.RemoveEmptyEntries));
        return nome.Replace(" ", "_").ToUpperInvariant();
    }
}
