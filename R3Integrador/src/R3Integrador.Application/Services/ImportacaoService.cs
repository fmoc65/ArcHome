using Microsoft.Extensions.Logging;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Mappers;

namespace R3Integrador.Application.Services;

public class ImportacaoService
{
    private const string PastaSaidaPadrao = @"C:\Projetos\ArcHome\R3Integrador\Saida";

    private readonly IExcelReader _excelReader;
    private readonly IVinilicoReader _vinilicoReader;
    private readonly IDelcredereReader _delcredereReader;
    private readonly IVillaArtReader _villaArtReader;
    private readonly IExcelExporter _excelExporter;
    private readonly ILogger<ImportacaoService> _logger;

    public ImportacaoService(
        IExcelReader excelReader,
        IVinilicoReader vinilicoReader,
        IDelcredereReader delcredereReader,
        IVillaArtReader villaArtReader,
        IExcelExporter excelExporter,
        ILogger<ImportacaoService> logger)
    {
        _excelReader = excelReader;
        _vinilicoReader = vinilicoReader;
        _delcredereReader = delcredereReader;
        _villaArtReader = villaArtReader;
        _excelExporter = excelExporter;
        _logger = logger;
    }

    // 1. Processamento Padrão (Varejo, Vinilico, Lastra)
    public async Task ProcessarPadraoAsync(string caminhoArquivo, string tabela)
    {
        _logger.LogInformation("Iniciando a leitura do arquivo {Arquivo} na aba {Tabela}...", caminhoArquivo, tabela);

        // Lê os produtos normais
        var produtosBrutos = await _excelReader.LerAsync(caminhoArquivo, tabela);

        if (produtosBrutos == null || !produtosBrutos.Any())
        {
            _logger.LogWarning("Nenhum produto válido encontrado na aba {Tabela}.", tabela);
            return;
        }

        // ====================================================================
        // FIX DO ERRO CS1503: 
        // Converte a lista de 'ProdutoNormalizado' para 'ProdutoErpDto'
        // ====================================================================
        var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

        var arquivoSaida = CriarCaminhoSaida($"{tabela}_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

        // Agora o ExportarAsync recebe a lista convertida corretamente!
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);                

        _logger.LogInformation("Exportação da tabela {Tabela} concluída com sucesso.", tabela);
    }

    // 2. Processamento Específico para a Planilha Delcredere (Opção 4 do Menu )
   public async Task ProcessarDelcredereAsync(string caminhoArquivo)
{
    _logger.LogInformation("Iniciando processamento da tabela Delcredere Varejo...");

    // 1. Lê a planilha na Aba correspondente usando o seu ExcelReaderService
    var produtosBrutos = await _delcredereReader.LerAsync(caminhoArquivo); 

    if (produtosBrutos == null || !produtosBrutos.Any())
    {
        _logger.LogWarning("Nenhum produto válido encontrado na aba 'COM DEL CREDERE'.");
        return;
    }

    // 2. Transforma a lista através do Mapeador inteligente
    var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

    // 3. Define a pasta de saída padrão de forma dinâmica
    string arquivoSaida = CriarCaminhoSaida($"IMPORTACAO_ERP_DELCREDERE_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

    // 4. Exporta usando o novo padrão de 60 colunas do ERP
    await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

    _logger.LogInformation("Processamento e exportação Delcredere concluídos com sucesso.");
}

    // 3. Processamento Específico para a Planilha Villa Art Boutique (Opção 5 do Menu)
    public async Task ProcessarVillaArtAsync(string caminhoArquivo)
    {
        _logger.LogInformation("Iniciando processamento da tabela Villa Art Boutique...");
        
        var produtosBrutos = await _villaArtReader.LerAsync(caminhoArquivo);

        if (produtosBrutos == null || !produtosBrutos.Any()) return;

        var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

        var arquivoSaida = CriarCaminhoSaida($"VILLA_ART_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Exportação Villa Art concluída.");
    }

    // Processamento específico para a Planilha de Vinílicos (Opção 2 do Menu)
public async Task ProcessarVinilicoAsync(string caminhoArquivo)
{
    _logger.LogInformation("Iniciando processamento da tabela de Vinílicos...");

    // 1. Lê os dados da aba "VINILICO" da planilha do fornecedor
    var produtosBrutos = await _vinilicoReader.LerAsync(caminhoArquivo);

    if (produtosBrutos == null || !produtosBrutos.Any())
    {
        _logger.LogWarning("Nenhum produto válido encontrado na aba VINILICO.");
        return;
    }

    // 2. Mapeia os dados normalizados para o DTO de 60 colunas do ERP usando o ProdutoErpMapper
    var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

    // 3. Define o caminho final de saída para a pasta unificada
    var arquivoSaida = CriarCaminhoSaida($"VINILICO_ERP_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
    
    // 4. Exporta usando o ExcelExportService que já possui as 60 colunas estruturadas
    await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

    _logger.LogInformation("Processamento e exportação de Vinílicos concluídos com sucesso.");
}

private static string CriarCaminhoSaida(string nomeArquivo)
{
    Directory.CreateDirectory(PastaSaidaPadrao);
    return Path.Combine(PastaSaidaPadrao, nomeArquivo);
}
}
