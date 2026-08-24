using System.Globalization;

namespace R3Integrador.Application.Services;

/// <summary>
/// Premissas fiscais fornecidas para a tabela Nina Martinelli.
/// As alíquotas são armazenadas como frações decimais: 12% = 0.12m.
/// </summary>
public sealed record NinaFiscalParameters
{
    public string RegimeEmpresa { get; init; } = "SIMPLES NACIONAL";
    public string UfEmpresa { get; init; } = "SP";
    public string UfFabrica { get; init; } = "SP";
    public decimal AliquotaIcmsOrigem { get; init; } = 0.12m;
    public decimal AliquotaIcmsSaida { get; init; } = 0.12m;
    public decimal Mva { get; init; } = 0.81m;
    public decimal Ipi { get; init; } = 0.0065m;
    public decimal PisOrigem { get; init; } = 0.0165m;
    public decimal CofinsOrigem { get; init; } = 0.076m;
    public decimal PercentualSt { get; init; } = 0.0986m;

    public string PisOrigemTexto => PisOrigem.ToString("0.####", CultureInfo.InvariantCulture);
    public string CofinsOrigemTexto => CofinsOrigem.ToString("0.####", CultureInfo.InvariantCulture);
}

