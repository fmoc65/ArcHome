namespace R3Integrador.Application.Services;

public interface ICsosnService
{
    CsosnResultado Determinar(CsosnCenario cenario);
}

public sealed record CsosnCenario(
    decimal PercentualStInformado,
    bool SemIncidenciaSt,
    bool IcmsStCobradoAnteriormente = false,
    bool PermiteCredito = false,
    bool OperacaoNaoTributada = false);

public sealed record CsosnResultado(string Codigo, decimal PercentualSt);

