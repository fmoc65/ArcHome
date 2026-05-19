using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IExcelExporter
{
    Task ExportarAsync(List<ProdutoErpDto> produtos, string caminhoSaida);
}