using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging; // Necessário para ClearProviders
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Services;
using R3Integrador.Infrastructure.Export;
using R3Integrador.Infrastructure.Persistence;
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
builder.Services.AddSingleton(new NinaFiscalParameters());
builder.Services.AddSingleton<ICsosnService, CsosnService>();
builder.Services.AddSingleton<IVinilicoReader, VinilicoReaderService>();
builder.Services.AddSingleton<IDelcredereReader, DelcredereReaderService>();
builder.Services.AddSingleton<IVillaArtReader, VillaArtReaderService>();
builder.Services.AddSingleton<ILastraReader, LastraReaderService>();
builder.Services.AddSingleton<IRubinettosReader, RubinettosReaderService>();
builder.Services.AddSingleton<IRocaReader, RocaReaderService>();
builder.Services.AddSingleton<IImersiReader, ImersiReaderService>();
builder.Services.AddSingleton<IStudioMorandinReader, StudioMorandinReaderService>();
builder.Services.AddSingleton<IAdamaReader, AdamaReaderService>();
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
    AdamaReader = sp.GetRequiredService<IAdamaReader>(),
    InvitaReader = sp.GetRequiredService<IInvitaReader>(),
    DerossoReader = sp.GetRequiredService<IDerossoReader>(),
    AtlasReader = sp.GetRequiredService<IAtlasReader>(),
    NinaMartinelliReader = sp.GetRequiredService<INinaMartinelliReader>(),
    SpecialSlReader = sp.GetRequiredService<ISpecialSlReader>()
});
builder.Services.AddSingleton<ImportacaoService>();
builder.Services.AddSingleton<IExcelExporter, ExcelExportService>();
builder.Services.AddSingleton<IExcelResumidoExporter, ExcelResumidoExportService>();
builder.Services.AddSingleton<RevisaoCamposFiscaisService>();
builder.Services.AddSingleton<R3IntegradorDbInitializer>();
builder.Services.AddSingleton<AtualizacaoPrecoFabricaService>();
builder.Services.AddSingleton<VillagresDelcredereAtualizacaoService>();
builder.Services.AddSingleton<AdamaAtualizacaoService>();

var app = builder.Build();

if (args.Length > 0 && args[0].Equals("--processar-adama-com-banco", StringComparison.OrdinalIgnoreCase))
{
    if (args.Length is < 3 or > 5)
    {
        Console.Error.WriteLine(
            "Uso: --processar-adama-com-banco <tabela-fornecedor.xlsx> <contador-ok.xlsx> [R3IntegradorDb.db] [pasta-saida]");
        return;
    }

    var caminhoBanco = args.Length >= 4
        ? args[3]
        : Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db");
    var pastaSaida = args.Length >= 5
        ? args[4]
        : Path.Combine(Directory.GetCurrentDirectory(), "Saida");
    var resultado = await app.Services.GetRequiredService<AdamaAtualizacaoService>()
        .ExecutarAsync(args[1], args[2], caminhoBanco, pastaSaida);

    Console.WriteLine($"[OK] Produtos Adama validados: {resultado.ProdutosValidadosNoBanco}");
    Console.WriteLine($"[OK] Produtos com preco corrigido: {resultado.ProdutosComPrecoCorrigido}");
    Console.WriteLine($"[OK] Produtos com unidade corrigida: {resultado.ProdutosComUnidadeCorrigida}");
    Console.WriteLine($"[OK] Atualizacao resumida dos produtos ja importados: {resultado.CaminhoPlanilhaAtualizacao}");
    Console.WriteLine($"[OK] Base completa reservada para auditoria: {resultado.CaminhoPlanilhaCorrigida}");
    Console.WriteLine($"[OK] Banco: {resultado.CaminhoBanco}");
    Console.WriteLine($"[OK] Backup do banco: {resultado.CaminhoBackupBanco}");
    Console.WriteLine($"[OK] Auditoria: {resultado.CaminhoAuditoria}");
    return;
}

if (args.Length > 0 && args[0].Equals("--atualizar-preco-fabrica", StringComparison.OrdinalIgnoreCase))
{
    if (args.Length is < 2 or > 4)
    {
        Console.Error.WriteLine(
            "Uso: --atualizar-preco-fabrica <planilha.xlsx> [R3IntegradorDb.db] [pasta-saida]");
        return;
    }

    var caminhoPlanilha = args[1];
    var caminhoBanco = args.Length >= 3
        ? args[2]
        : Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db");
    var pastaSaida = args.Length >= 4
        ? args[3]
        : Path.Combine(Directory.GetCurrentDirectory(), "Saida");

    var resultado = await app.Services.GetRequiredService<AtualizacaoPrecoFabricaService>()
        .ExecutarAsync(caminhoBanco, caminhoPlanilha, pastaSaida);

    Console.WriteLine($"[OK] Precos atualizados no banco: {resultado.ProdutosAtualizadosNoBanco}");
    Console.WriteLine($"[OK] Planilha de atualizacao: {resultado.CaminhoPlanilhaAtualizacao}");
    Console.WriteLine($"[OK] Planilha de inclusao: {resultado.CaminhoPlanilhaInclusao}");
    Console.WriteLine($"[OK] Referencias para inclusao: {string.Join(", ", resultado.ReferenciasParaInclusao)}");
    return;
}

if (args.Length > 0 && args[0].Equals("--criar-banco", StringComparison.OrdinalIgnoreCase))
{
    if (args.Length is < 2 or > 4)
    {
        Console.Error.WriteLine(
            "Uso: --criar-banco <planilha.xlsx> [R3IntegradorDb.db] [tabela-origem]");
        return;
    }

    var caminhoPlanilha = args[1];
    var caminhoBanco = args.Length >= 3
        ? args[2]
        : Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db");
    var tabelaOrigem = args.Length >= 4
        ? args[3]
        : InferirTabelaOrigem(caminhoPlanilha);

    var resultado = app.Services.GetRequiredService<R3IntegradorDbInitializer>()
        .CriarOuAtualizar(caminhoBanco, caminhoPlanilha, tabelaOrigem);

    Console.WriteLine($"[OK] Banco: {resultado.CaminhoBanco}");
    Console.WriteLine($"[OK] Origem: {resultado.TabelaOrigem} (aba {resultado.AbaPlanilha})");
    Console.WriteLine($"[OK] Produtos importados: {resultado.ProdutosImportados}");
    return;
}

if (args.Length > 0 && args[0].Equals("--atualizar-delcredere-com-banco", StringComparison.OrdinalIgnoreCase))
{
    if (args.Length is < 2 or > 4)
    {
        Console.Error.WriteLine(
            "Uso: --atualizar-delcredere-com-banco <tabela.xlsx> [R3IntegradorDb.db] [pasta-saida]");
        return;
    }

    var caminhoPlanilha = args[1];
    var caminhoBanco = args.Length >= 3
        ? args[2]
        : Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db");
    var pastaSaida = args.Length >= 4
        ? args[3]
        : Path.Combine(Directory.GetCurrentDirectory(), "Saida");
    var resultado = await app.Services.GetRequiredService<VillagresDelcredereAtualizacaoService>()
        .ExecutarAsync(caminhoPlanilha, caminhoBanco, pastaSaida);

    Console.WriteLine($"[OK] Pasta de saida: {resultado.PastaSaida}");
    Console.WriteLine($"[OK] Referencias para atualizacao: {resultado.ReferenciasAtualizacao}");
    Console.WriteLine($"[OK] Referencias para inclusao: {resultado.ReferenciasInclusao}");
    Console.WriteLine($"[OK] Referencias do banco sem preco na nova tabela: {resultado.ReferenciasBancoAusentesNaTabela}");
    Console.WriteLine($"[OK] Auditoria: {resultado.CaminhoAuditoria}");
    return;
}

if (args.Length > 0 && args[0].Equals("--sincronizar-confirmados-erp", StringComparison.OrdinalIgnoreCase))
{
    if (args.Length is < 2 or > 4)
    {
        Console.Error.WriteLine(
            "Uso: --sincronizar-confirmados-erp <tabela.xlsx> [R3IntegradorDb.db] [pasta-saida]");
        return;
    }

    var caminhoPlanilha = args[1];
    var caminhoBanco = args.Length >= 3
        ? args[2]
        : Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db");
    var pastaSaida = args.Length >= 4
        ? args[3]
        : Path.Combine(Directory.GetCurrentDirectory(), "Saida");
    var resultado = await app.Services.GetRequiredService<VillagresDelcredereAtualizacaoService>()
        .SincronizarConfirmadosNoErpAsync(caminhoPlanilha, caminhoBanco, pastaSaida);

    Console.WriteLine($"[OK] Produtos sincronizados na replica local: {resultado.ProdutosInseridos}");
    Console.WriteLine($"[OK] Backup: {resultado.CaminhoBackup}");
    Console.WriteLine($"[OK] Referencias: {string.Join(", ", resultado.ReferenciasInseridas)}");
    return;
}

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
    Console.WriteLine("   18 - Processar Planilha ADAMA (PROVISORIA)");
    Console.WriteLine();
    Console.WriteLine("  [REVISAO FISCAL]");
    Console.WriteLine("   15 - Sinalizar campos fiscais obrigatorios ausentes");
    Console.WriteLine();
    Console.WriteLine("  [ATUALIZAÇÕES DE PREÇO]");
    Console.WriteLine("   16 - Atualização Del Credere Villagres (layout resumido)");
    Console.WriteLine("   17 - Atualização e inclusão Del Credere Villagres pelo banco");
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
        case "18": await ExecutarAcaoAsync(() => importacaoService.ProcessarAdamaAsync(ObterCaminho("Digite ou arraste a tabela ADAMA:"))); break;
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
        case "16":
            await ExecutarAcaoAsync(() => importacaoService.ProcessarAtualizacaoDelcredereAsync(
                ObterCaminho("Digite ou arraste a TABELA DELCREDERE VAREJO:")));
            break;
        case "17":
        {
            var caminhoTabela = ObterCaminho("Digite ou arraste a TABELA DELCREDERE VAREJO:");
            var resultado = await app.Services.GetRequiredService<VillagresDelcredereAtualizacaoService>()
                .ExecutarAsync(
                    caminhoTabela,
                    Path.Combine(Directory.GetCurrentDirectory(), "R3IntegradorDb.db"),
                    Path.Combine(Directory.GetCurrentDirectory(), "Saida"));
            Console.WriteLine($"[OK] Atualizacao: {resultado.ReferenciasAtualizacao}; inclusao: {resultado.ReferenciasInclusao}.");
            Console.WriteLine($"[OK] Auditoria: {resultado.CaminhoAuditoria}");
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

static string InferirTabelaOrigem(string caminhoPlanilha)
{
    var nomeArquivo = Path.GetFileNameWithoutExtension(caminhoPlanilha).ToUpperInvariant();
    foreach (var origem in new[] { "VAREJO", "VINILICO", "LASTRA" })
    {
        if (nomeArquivo.Contains(origem, StringComparison.Ordinal))
        {
            return origem;
        }
    }

    return "IMPORTACAO_ERP";
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
