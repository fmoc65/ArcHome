using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging; // Necessário para ClearProviders
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Services;
using R3Integrador.Infrastructure.Export;
using R3Integrador.Infrastructure.Repositories;
using Serilog;

var builder = Host.CreateApplicationBuilder(args);
builder.Configuration.AddJsonFile(
    Path.Combine(AppContext.BaseDirectory, "appsettings.json"),
    optional: true,
    reloadOnChange: false);

Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .CreateLogger();

// VÍNCULO DO LOG: Limpa os provedores padrão e injeta o Serilog no Container do Framework
builder.Logging.ClearProviders();
builder.Logging.AddSerilog();

builder.Services.AddSingleton<IExcelReader, ExcelReaderService>();
builder.Services.AddSingleton<IVinilicoReader, VinilicoReaderService>();
builder.Services.AddSingleton<IDelcredereReader, DelcredereReaderService>();
builder.Services.AddSingleton<IVillaArtReader, VillaArtReaderService>();
builder.Services.AddSingleton<ILastraReader, LastraReaderService>();
builder.Services.AddSingleton<IRubinettosReader, RubinettosReaderService>();
builder.Services.AddSingleton(sp => new ImportacaoReaderSet
{
    ExcelReader = sp.GetRequiredService<IExcelReader>(),
    VinilicoReader = sp.GetRequiredService<IVinilicoReader>(),
    DelcredereReader = sp.GetRequiredService<IDelcredereReader>(),
    VillaArtReader = sp.GetRequiredService<IVillaArtReader>(),
    LastraReader = sp.GetRequiredService<ILastraReader>(),
    RubinettosReader = sp.GetRequiredService<IRubinettosReader>()
});
builder.Services.AddSingleton<ImportacaoService>();
builder.Services.AddSingleton<IExcelExporter, ExcelExportService>();

var app = builder.Build();

Log.Information("Sistema integrado e motores de log prontos para uso.");
bool executando = true;

// ... Bloco de configuração inicial do appsettings / Serilog igual ao anterior ...

while (executando)
{
    if (!Console.IsInputRedirected)
    {
        Console.Clear();
    }

    Console.ForegroundColor = ConsoleColor.Cyan;
    Console.WriteLine("=================================================================");
    Console.WriteLine("             R3 INTEGRADOR - COMPATIBILIZADOR DE FORNECEDORES    ");
    Console.WriteLine("=================================================================");
    Console.ResetColor();
    Console.WriteLine();
    Console.WriteLine("  [FORNECEDOR PADRÃO - POR ABAS]");
    Console.WriteLine("    1 - Processar Aba VAREJO");
    Console.WriteLine("    2 - Processar Aba VINILICO");
    Console.WriteLine("    3 - Processar Aba LASTRA");
    Console.WriteLine();
    Console.WriteLine("  [FORNECEDORES ESPECÍFICOS]");
    Console.WriteLine("    4 - Processar Planilha TABELA DELCREDERE VAREJO");
    Console.WriteLine("    5 - Processar Planilha VILLA ART - BOUTIQUE");
    Console.WriteLine("    6 - Processar Planilha RUBINETTOS");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("    0 - Sair");
    Console.ResetColor();
    Console.WriteLine();
    Console.Write("Escolha a opção correspondente ao arquivo que deseja importar: ");

    var opcao = Console.ReadLine();
    var importacaoService = app.Services.GetRequiredService<ImportacaoService>();

    switch (opcao)
    {
        case "1": await ExecutarAcaoAsync(() => importacaoService.ProcessarPadraoAsync(ObterCaminho(), "VAREJO")); break;
        case "2": await ExecutarAcaoAsync(() => importacaoService.ProcessarVinilicoAsync(ObterCaminho())); break;
        case "3": await ExecutarAcaoAsync(() => importacaoService.ProcessarLastraAsync(ObterCaminho())); break;
        case "4": await ExecutarAcaoAsync(() => importacaoService.ProcessarDelcredereAsync(ObterCaminho())); break;
        case "5": await ExecutarAcaoAsync(() => importacaoService.ProcessarVillaArtAsync(ObterCaminho())); break;
        case "6": await ExecutarAcaoAsync(() => importacaoService.ProcessarRubinettosAsync(ObterCaminho())); break;
        case "0": executando = false; break;
        default: Console.WriteLine("Opção inválida."); break;
    }

    if (executando)
    {
        Console.WriteLine("\nPressione ENTER para voltar ao menu...");
        Console.ReadLine();
    }
}

static string ObterCaminho()
{
    Console.Write("\nDigite ou arraste o arquivo Excel do Fornecedor: ");
    var caminho = Console.ReadLine() ?? string.Empty;
    return caminho.Trim('"');
}

static async Task ExecutarAcaoAsync(Func<Task> acao)
{
    try
    {
        await acao();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("\n[OK] Processamento finalizado!");
        Console.ResetColor();
    }
    catch (Exception ex)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.WriteLine($"\n[ERRO CRÍTICO] Falha no fluxo: {ex.Message}");
        Console.ResetColor();
    }
}
