using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IAdamaReader
{
    Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo);
}
