namespace R3R3Integrador.Application.DTOs;

public class ProdutoNormalizado
{
    public string TipoTabela { get; set; } = string.Empty;
    public string Formato { get; set; } = string.Empty;
    public string Referencia { get; set; } = string.Empty;
    public string Linha { get; set; } = string.Empty;
    public string Colecao { get; set; } = string.Empty;
    public string Cor { get; set; } = string.Empty;
    public string Superficie { get; set; } = string.Empty;
    public int Faces { get; set; }
    public string VariacaoTonalidade { get; set; } = string.Empty;
    public decimal PrecoTabela { get; set; }
    public decimal PrecoDesconto { get; set; }
    public decimal Espessura { get; set; }
    public decimal M2Caixa { get; set; }
}