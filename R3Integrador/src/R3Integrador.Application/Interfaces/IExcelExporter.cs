using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IExcelExporter
{
    Task ExportarAsync(List<ProdutoNormalizado> produtos,  string caminhoSaida);
    
}