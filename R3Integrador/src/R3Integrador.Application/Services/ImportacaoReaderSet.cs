using R3Integrador.Application.Interfaces;

namespace R3Integrador.Application.Services;

public sealed class ImportacaoReaderSet
{
    public required IExcelReader ExcelReader { get; init; }
    public required IVinilicoReader VinilicoReader { get; init; }
    public required IDelcredereReader DelcredereReader { get; init; }
    public required IVillaArtReader VillaArtReader { get; init; }
    public required ILastraReader LastraReader { get; init; }
    public required IRubinettosReader RubinettosReader { get; init; }
}
