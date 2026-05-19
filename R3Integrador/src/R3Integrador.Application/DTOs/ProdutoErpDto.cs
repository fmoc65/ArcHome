namespace R3Integrador.Application.DTOs;

public class ProdutoErpDto
{
    public string CodigoInterno { get; set; } = string.Empty;
    public string CodigoFabrica { get; set; } = string.Empty;
    public string CodigoBarras { get; set; } = string.Empty;
    public string DescricaoCompleta { get; set; } = string.Empty;
    public string DescricaoComercial { get; set; } = string.Empty;
    public string Grupo { get; set; } = string.Empty;
    public string SubGrupo { get; set; } = string.Empty;
    public string Marca { get; set; } = string.Empty;
    public string Linha { get; set; } = string.Empty;
    public string Modelo { get; set; } = string.Empty;
    public string Cor { get; set; } = string.Empty;
    public string Unidade { get; set; } = "CX";
    public decimal PrecoVenda { get; set; }
    public decimal PrecoFabrica { get; set; }
    public string Ncm { get; set; } = string.Empty;
    public string Cst { get; set; } = string.Empty;
    public string CfopDentro { get; set; } = string.Empty;
    public string CfopFora { get; set; } = string.Empty;
    public string Cfop { get; set; } = string.Empty;
    public decimal PrecoFracionado { get; set; }
    public decimal PesoLiquido { get; set; }
    public decimal PesoBruto { get; set; }
    public decimal Espessura { get; set; }
    public string Observacao { get; set; } = string.Empty;
}