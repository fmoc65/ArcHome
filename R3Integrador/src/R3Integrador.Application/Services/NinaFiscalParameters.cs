using System.Globalization;

namespace R3Integrador.Application.Services;

/// <summary>
/// Premissas fiscais confirmadas por Mônica para a importação Nina Martinelli.
/// O ERP recebe as alíquotas em pontos percentuais: 12% = 12m.
/// </summary>
public sealed record NinaFiscalParameters
{
    public string NcmProdutosCeramicos { get; init; } = "69072300";
    public string NcmProdutosCimenticios { get; init; } = "68101900";
    public string NcmImpermeabilizantes { get; init; } = "38099390";
    public string RegimeEmpresa { get; init; } = "SIMPLES NACIONAL";
    public string UfEmpresa { get; init; } = "SP";
    public string UfFabrica { get; init; } = "SP";
    public decimal AliquotaIcmsOrigem { get; init; } = 12m;
    public decimal AliquotaIcmsSaida { get; init; } = 12m;
    public decimal Mva { get; init; } = 81m;
    public decimal Ipi { get; init; } = 0.65m;
    public decimal PisOrigem { get; init; } = 1.65m;
    public decimal CofinsOrigem { get; init; } = 7.60m;
    public decimal PercentualSt { get; init; } = 0m;
    public string Cst { get; init; } = "000";
    public string AliquotaCofinsCst { get; init; } = "99";
    public string AliquotaIpiCst { get; init; } = "99";
    public string AliquotaPisCst { get; init; } = "99";
    public string Csosn { get; init; } = "500";
    public string CfopDentro { get; init; } = "5405";
    public string CfopFora { get; init; } = "6404";
    public string EnquadramentoIpi { get; init; } = "999";
    public string AliquotaIbs { get; init; } = "0.1";
    public string AliquotaCbs { get; init; } = "0.9";
    public string ClassificacaoTributaria { get; init; } = "000001";

    public string PisOrigemTexto => PisOrigem.ToString("0.00", CultureInfo.InvariantCulture);
    public string CofinsOrigemTexto => CofinsOrigem.ToString("0.00", CultureInfo.InvariantCulture);
}
