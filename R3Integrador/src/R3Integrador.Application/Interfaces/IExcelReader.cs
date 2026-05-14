using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IExcelReader
{
    Task<List<ProdutoNormalizado>> LerAsync(
        string caminhoArquivo,
        string tabela);
}