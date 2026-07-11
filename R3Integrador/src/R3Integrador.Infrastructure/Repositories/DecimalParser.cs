using System.Globalization;

namespace R3Integrador.Infrastructure.Repositories;

internal static class DecimalParser
{
    public static decimal Parse(string valor)
    {
        if (string.IsNullOrWhiteSpace(valor))
        {
            return 0;
        }

        valor = valor.Replace("R$", string.Empty)
            .Replace("%", string.Empty)
            .Replace("\u00A0", string.Empty)
            .Trim();

        if (valor == "-")
        {
            return 0;
        }

        valor = valor.Replace(" ", string.Empty);

        var ultimoPonto = valor.LastIndexOf('.');
        var ultimaVirgula = valor.LastIndexOf(',');

        if (ultimoPonto >= 0 && ultimaVirgula >= 0)
        {
            valor = ultimaVirgula > ultimoPonto
                ? valor.Replace(".", string.Empty).Replace(",", ".")
                : valor.Replace(",", string.Empty);
        }
        else if (ultimaVirgula >= 0)
        {
            valor = valor.Replace(",", ".");
        }
        else if (valor.Count(c => c == '.') > 1)
        {
            var ultimoSeparador = valor.LastIndexOf('.');
            valor = valor[..ultimoSeparador].Replace(".", string.Empty) + valor[ultimoSeparador..];
        }

        return decimal.TryParse(
            valor,
            NumberStyles.Number | NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture,
            out var resultado)
            ? resultado
            : 0;
    }
}
