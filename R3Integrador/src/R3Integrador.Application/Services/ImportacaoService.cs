using R3Integrador.Application.Interfaces;

namespace R3Integrador.Application.Services;
public class ImportacaoService
{
    private readonly IExcelReader _excelReader;
    private readonly IExcelExporter _excelExporter;

    public ImportacaoService(
        IExcelReader excelReader,
        IExcelExporter excelExporter)
    {
        _excelReader = excelReader;
        _excelExporter = excelExporter;
    }

    public async Task ProcessarAsync(
        string caminhoArquivo,
        string tabela)
    {
        var produtos =
            await _excelReader.LerAsync(
                caminhoArquivo,
                tabela);

        var arquivoSaida = $@"C:\Projetos\ArcHome\R3Integrador\Saida\{tabela}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";

        await _excelExporter.ExportarAsync(
            produtos,
            arquivoSaida);                

        // regras futuras
    }
}