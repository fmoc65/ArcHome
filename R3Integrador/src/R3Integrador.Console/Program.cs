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
builder.Services.AddSingleton<IRocaReader, RocaReaderService>();
builder.Services.AddSingleton<IImersiReader, ImersiReaderService>();
builder.Services.AddSingleton<IStudioMorandinReader, StudioMorandinReaderService>();
builder.Services.AddSingleton<IInvitaReader, InvitaReaderService>();
builder.Services.AddSingleton<IDerossoReader, DerossoReaderService>();
builder.Services.AddSingleton<IAtlasReader, AtlasReaderService>();
builder.Services.AddSingleton<INinaMartinelliReader, NinaMartinelliReaderService>();
builder.Services.AddSingleton<ISpecialSlReader, SpecialSlReaderService>();
builder.Services.AddSingleton(sp => new ImportacaoReaderSet
{
    ExcelReader = sp.GetRequiredService<IExcelReader>(),
    VinilicoReader = sp.GetRequiredService<IVinilicoReader>(),
    DelcredereReader = sp.GetRequiredService<IDelcredereReader>(),
    VillaArtReader = sp.GetRequiredService<IVillaArtReader>(),
    LastraReader = sp.GetRequiredService<ILastraReader>(),
    RubinettosReader = sp.GetRequiredService<IRubinettosReader>(),
    RocaReader = sp.GetRequiredService<IRocaReader>(),
    ImersiReader = sp.GetRequiredService<IImersiReader>(),
    StudioMorandinReader = sp.GetRequiredService<IStudioMorandinReader>(),
    InvitaReader = sp.GetRequiredService<IInvitaReader>(),
    DerossoReader = sp.GetRequiredService<IDerossoReader>(),
    AtlasReader = sp.GetRequiredService<IAtlasReader>(),
    NinaMartinelliReader = sp.GetRequiredService<INinaMartinelliReader>(),
    SpecialSlReader = sp.GetRequiredService<ISpecialSlReader>()
});
builder.Services.AddSingleton<ImportacaoService>();
builder.Services.AddSingleton<IExcelExporter, ExcelExportService>();
builder.Services.AddSingleton<RevisaoCamposFiscaisService>();

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
    Console.WriteLine("    7 - Processar Planilha ROCA");
    Console.WriteLine("    8 - Processar Planilha IMERSI");
    Console.WriteLine("    9 - Processar Planilha STUDIO MORANDIN (PROVISORIA)");
    Console.WriteLine("   10 - Processar Planilha INVITA (PROVISORIA)");
    Console.WriteLine("   11 - Processar Planilha DEROSSO (PROVISORIA)");
    Console.WriteLine("   12 - Processar Planilha ATLAS REVENDA 35% (PROVISORIA)");
    Console.WriteLine("   13 - Processar Planilha NINA MARTINELLI (PROVISORIA)");
    Console.WriteLine("   14 - Processar TABELA ESPECIAL SL");
    Console.WriteLine("   15 - Sinalizar campos fiscais obrigatorios ausentes");
    Console.ForegroundColor = ConsoleColor.Red;
    Console.WriteLine("    0 - Sair");
    Console.ResetColor();
    Console.WriteLine();
    Console.Write("Escolha a opção correspondente ao arquivo que deseja importar: ");

    var opcao = Console.ReadLine();
    var importacaoService = app.Services.GetRequiredService<ImportacaoService>();

    switch (opcao)
    {
        case "1": await ExecutarAcaoAsync(() => importacaoService.ProcessarPadraoAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"), "VAREJO")); break;
        case "2": await ExecutarAcaoAsync(() => importacaoService.ProcessarVinilicoAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "3": await ExecutarAcaoAsync(() => importacaoService.ProcessarLastraAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "4":
        {
            var caminhoTabela = ObterCaminho("Digite o caminho da planilha TABELA DELCREDERE VAREJO:");
            var caminhoAliquotas = ObterCaminho("Digite o caminho da planilha IMPORTACAO_ERP_DELCREDERE_DEL20 (opcional):");
            await ExecutarAcaoAsync(() => importacaoService.ProcessarDelcredereAsync(caminhoTabela, string.IsNullOrWhiteSpace(caminhoAliquotas) ? null : caminhoAliquotas));
            break;
        }
        case "5": await ExecutarAcaoAsync(() => importacaoService.ProcessarVillaArtAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "6": await ExecutarAcaoAsync(() => importacaoService.ProcessarRubinettosAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "7": await ExecutarAcaoAsync(() => importacaoService.ProcessarRocaAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "8": await ExecutarAcaoAsync(() => importacaoService.ProcessarImersiAsync(ObterCaminho("Digite ou arraste o arquivo Excel do Fornecedor:"))); break;
        case "9": await ExecutarAcaoAsync(() => importacaoService.ProcessarStudioMorandinAsync(ObterCaminho("Digite ou arraste a tabela STUDIO MORANDIN:"))); break;
        case "10": await ExecutarAcaoAsync(() => importacaoService.ProcessarInvitaAsync(ObterCaminho("Digite ou arraste a tabela INVITA:"))); break;
        case "11": await ExecutarAcaoAsync(() => importacaoService.ProcessarDerossoAsync(ObterCaminho("Digite ou arraste a tabela DEROSSO:"))); break;
        case "12": await ExecutarAcaoAsync(() => importacaoService.ProcessarAtlasAsync(ObterCaminho("Digite ou arraste a tabela ATLAS:"))); break;
        case "13": await ExecutarAcaoAsync(() => importacaoService.ProcessarNinaMartinelliAsync(ObterCaminho("Digite ou arraste a tabela NINA MARTINELLI convertida para XLSX:"))); break;
        case "14": await ExecutarAcaoAsync(() => importacaoService.ProcessarSpecialSlAsync(ObterCaminho("Digite ou arraste o PDF ou XLSX TABELA ESPECIAL SL:"))); break;
        case "15":
        {
            var resultado = app.Services.GetRequiredService<RevisaoCamposFiscaisService>()
                .Revisar(ObterCaminho("Digite ou arraste a importacao ERP a revisar:"));
            Console.WriteLine($"[OK] Copia revisada gerada em: {resultado.CaminhoSaida}");
            foreach (var (campo, faltas) in resultado.FaltasPorCampo.Where(item => item.Value > 0))
            {
                Console.WriteLine($"  - {campo}: {faltas} item(ns) sem preenchimento");
            }
            break;
        }
        case "0": executando = false; break;
        default: Console.WriteLine("Opção inválida."); break;
    }

    if (executando)
    {
        Console.WriteLine("\nPressione ENTER para voltar ao menu...");
        Console.ReadLine();
    }
}

static string ObterCaminho(string prompt)
{
    Console.Write($"\n{prompt} ");
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
