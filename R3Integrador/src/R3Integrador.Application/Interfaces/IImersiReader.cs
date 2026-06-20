using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Interfaces;

public interface IImersiReader
{
    Task<Dictionary<string, List<ProdutoErpDto>>> LerAsync(string caminhoArquivo);
}
