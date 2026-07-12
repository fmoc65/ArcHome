using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IStudioMorandinReader
{
    Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo);
}
