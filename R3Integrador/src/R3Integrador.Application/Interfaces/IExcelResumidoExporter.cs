using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IExcelResumidoExporter
{
    Task ExportarAsync(IReadOnlyCollection<ProdutoErpDto> produtos, string caminhoSaida);
}
