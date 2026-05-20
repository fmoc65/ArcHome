namespace R3Integrador.Application.DTOs;

public class ProdutoErpDto
{
    // Colunas 1 a 10
    public string CodigoInterno { get; set; } = string.Empty;       // Coluna 1
    public string CodigoFabrica { get; set; } = string.Empty;       // Coluna 2
    public string CodigoBarras { get; set; } = string.Empty;        // Coluna 3
    public string DescricaoCompleta { get; set; } = string.Empty;   // Coluna 4
    public string DescricaoComercial { get; set; } = string.Empty;  // Coluna 5
    public string Grupo { get; set; } = string.Empty;               // Coluna 6
    public string SubGrupo { get; set; } = string.Empty;            // Coluna 7
    public string Marca { get; set; } = string.Empty;               // Coluna 8
    public string Linha { get; set; } = string.Empty;               // Coluna 9
    public string Modelo { get; set; } = string.Empty;              // Coluna 10

    // Colunas 11 a 20
    public string Voltagem { get; set; } = "N/A";                    // Coluna 11
    public string Cor { get; set; } = string.Empty;                 // Coluna 12
    public string Ncm { get; set; } = string.Empty;                 // Coluna 13
    public string UfOrigem { get; set; } = "SP";                    // Coluna 14
    public decimal PrecoVenda { get; set; }                         // Coluna 15
    public decimal PrecoFabrica { get; set; }                       // Coluna 16
    public decimal DescontoPercentual { get; set; }                 // Coluna 17
    public decimal IpiPercentual { get; set; }                      // Coluna 18
    public decimal AliqIcmsOrigem { get; set; }                     // Coluna 19
    public decimal AliqIcmsInterna { get; set; }                    // Coluna 20

    // Colunas 21 a 30
    public decimal Iva { get; set; }                                // Coluna 21
    public decimal FreteReais { get; set; }                         // Coluna 22
    public decimal FretePercentual { get; set; }                    // Coluna 23
    public string Unidade { get; set; } = "CX";                     // Coluna 24
    public decimal QtdeEmbalagemVenda { get; set; } = 1;            // Coluna 25
    public string Cst { get; set; } = string.Empty;                 // Coluna 26
    public string AliquotaCofinsCst { get; set; } = string.Empty;   // Coluna 27
    public string AliquotaIpiCst { get; set; } = string.Empty;      // Coluna 28
    public string AliquotaPisCst { get; set; } = string.Empty;      // Coluna 29
    public string Csosn { get; set; } = string.Empty;               // Coluna 30

    // Colunas 31 a 40
    public string CfopDentro { get; set; } = string.Empty;          // Coluna 31
    public string CfopFora { get; set; } = string.Empty;            // Coluna 32
    public decimal PesoLiquido { get; set; }                        // Coluna 33
    public decimal PesoBruto { get; set; }                          // Coluna 34
    public decimal QtdeEmbalagemCompra { get; set; } = 1;           // Coluna 35
    public decimal ValorPi { get; set; }                            // Coluna 36
    public decimal AliquotaCofins { get; set; }                     // Coluna 37
    public decimal AliquotaPis { get; set; }                        // Coluna 38
    public decimal PercentualSt { get; set; }                       // Coluna 39
    public string UnidFabril { get; set; } = "CX";                  // Coluna 40

    // Colunas 41 a 50
    public string Observacao { get; set; } = string.Empty;          // Coluna 41
    public decimal DiferencaIcms { get; set; }                      // Coluna 42
    public decimal ReducaoBaseIcms { get; set; }                    // Coluna 43
    public decimal ReducaoBaseSt { get; set; }                      // Coluna 44
    public string RetencaoPis { get; set; } = string.Empty;         // Coluna 45
    public string RetencaoCofins { get; set; } = string.Empty;      // Coluna 46
    public string RetencaoCsll { get; set; } = string.Empty;        // Coluna 47
    public string RetencaoIrrf { get; set; } = string.Empty;        // Coluna 48
    public string RetencaoPrevSocial { get; set; } = string.Empty;  // Coluna 49
    public string Localizacao { get; set; } = string.Empty;         // Coluna 50

    // Colunas 51 a 60
    public string EnquadramentoIpi { get; set; } = string.Empty;    // Coluna 51
    public string AliquotaPisOrigem { get; set; } = string.Empty;   // Coluna 52
    public string AliquotaCofinsOrigem { get; set; } = string.Empty;// Coluna 53
    public string Imagem { get; set; } = string.Empty;              // Coluna 54
    public decimal EstoqueMinimo { get; set; }                      // Coluna 55
    public decimal EstoqueMaximo { get; set; }                      // Coluna 56
    public string AliquotaIbs { get; set; } = string.Empty;         // Coluna 57
    public string AliquotaCbs { get; set; } = string.Empty;         // Coluna 58
    public string ClassificacaoTributaria { get; set; } = string.Empty; // Coluna 59
    public string CodigoBeneficio { get; set; } = string.Empty;     // Coluna 60
}