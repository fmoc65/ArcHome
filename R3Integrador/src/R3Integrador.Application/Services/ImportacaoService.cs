using Microsoft.Extensions.Logging;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Mappers;

namespace R3Integrador.Application.Services;

public class ImportacaoService
{
    private readonly IExcelReader _excelReader;
    private readonly IExcelExporter _excelExporter;
    private readonly ILogger<ImportacaoService> _logger;

    public ImportacaoService(
        IExcelReader excelReader,
        IExcelExporter excelExporter,
        ILogger<ImportacaoService> logger)
    {
        _excelReader = excelReader;
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

        var pastaSaida = @"C:\Projetos\ArcHome\R3Integrador\Saida";
        if (!Directory.Exists(pastaSaida)) Directory.CreateDirectory(pastaSaida);

        var arquivoSaida = Path.Combine(pastaSaida, $"{tabela}_ERP_{DateTime.Now:yyyyMMddHHmmss}.xlsx");

        // Agora o ExportarAsync recebe a lista convertida corretamente!
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);                

        _logger.LogInformation("Exportação da tabela {Tabela} concluída com sucesso.", tabela);
    }

    // 2. Processamento Específico para a Planilha Delcredere (Opção 4 do Menu )
    public async Task ProcessarDelcredereAsync(string caminhoArquivo)
    {
        _logger.LogInformation("Iniciando processamento da tabela Delcredere...");
        
        // Exemplo: Reaproveitando a leitura padrão por enquanto (aba 1) ou crie um método específico no IExcelReader depois
        var produtosBrutos = await _excelReader.LerAsync(caminhoArquivo, "COM DEL CREDERE"); 

        if (produtosBrutos == null || !produtosBrutos.Any()) return;

        // Mapeia e já aplica uma regra de comissão, se necessário
        var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

        var arquivoSaida = $@"C:\Projetos\ArcHome\R3Integrador\Saida\DELCREDERE_ERP_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Exportação Delcredere concluída.");
    }

    // 3. Processamento Específico para a Planilha Villa Art Boutique (Opção 5 do Menu)
    public async Task ProcessarVillaArtAsync(string caminhoArquivo)
    {
        _logger.LogInformation("Iniciando processamento da tabela Villa Art Boutique...");
        
        var produtosBrutos = await _excelReader.LerAsync(caminhoArquivo, "VILLA ART - BOUTIQUE");

        if (produtosBrutos == null || !produtosBrutos.Any()) return;

        var produtosErp = produtosBrutos.Select(p => ProdutoErpMapper.Map(p)).ToList();

        var arquivoSaida = $@"C:\Projetos\ArcHome\R3Integrador\Saida\VILLA_ART_ERP_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
        await _excelExporter.ExportarAsync(produtosErp, arquivoSaida);

        _logger.LogInformation("Exportação Villa Art concluída.");
    }
}