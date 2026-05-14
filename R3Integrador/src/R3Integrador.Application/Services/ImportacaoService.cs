using R3Integrador.Application.Interfaces;

namespace R3Integrador.Application.Services;
public class ImportacaoService
{
    private readonly IExcelReader _excelReader;

    public ImportacaoService(
        IExcelReader excelReader)
    {
        _excelReader = excelReader;
    }

    public async Task ProcessarAsync(
        string caminhoArquivo,
        string tabela)
    {
        var produtos =
            await _excelReader.LerAsync(
                caminhoArquivo,
                tabela);

        // regras futuras
    }
}