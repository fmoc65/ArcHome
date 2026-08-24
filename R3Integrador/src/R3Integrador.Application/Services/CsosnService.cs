namespace R3Integrador.Application.Services;

/// <summary>
/// Determina o CSOSN a partir da situação efetiva da operação.
/// 101/201 são usados quando existe permissão de crédito; 102/202 quando não
/// existe; 400 para operação não tributada; e 500 somente quando o ICMS-ST já
/// foi cobrado anteriormente ou por antecipação.
/// </summary>
public sealed class CsosnService : ICsosnService
{
    private const decimal Tolerancia = 0.000001m;
    private readonly NinaFiscalParameters _parameters;

    public CsosnService(NinaFiscalParameters parameters)
    {
        _parameters = parameters;
    }

    public CsosnResultado Determinar(CsosnCenario cenario)
    {
        if (cenario.OperacaoNaoTributada)
        {
            return new CsosnResultado("400", 0m);
        }

        if (cenario.IcmsStCobradoAnteriormente)
        {
            return new CsosnResultado("500", _parameters.PercentualSt);
        }

        if (cenario.SemIncidenciaSt || cenario.PercentualStInformado == 0m)
        {
            return new CsosnResultado(cenario.PermiteCredito ? "101" : "102", 0m);
        }

        var percentualNormalizado = cenario.PercentualStInformado > 1m
            ? cenario.PercentualStInformado / 100m
            : cenario.PercentualStInformado;

        if (Math.Abs(percentualNormalizado - _parameters.PercentualSt) > Tolerancia)
        {
            throw new InvalidOperationException(
                $"Percentual de ST {cenario.PercentualStInformado} inválido. " +
                $"Esperado: {_parameters.PercentualSt} (9,86%).");
        }

        return new CsosnResultado(
            cenario.PermiteCredito ? "201" : "202",
            _parameters.PercentualSt);
    }
}

