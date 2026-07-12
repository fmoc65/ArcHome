using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface INinaMartinelliReader
{
    Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo);
}
