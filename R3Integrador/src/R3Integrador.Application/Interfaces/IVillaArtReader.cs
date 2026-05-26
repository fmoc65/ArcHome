using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IVillaArtReader
{
    Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo);
}
