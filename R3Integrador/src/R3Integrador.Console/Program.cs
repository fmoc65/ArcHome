using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Services;
using R3Integrador.Infrastructure.Export;
using R3Integrador.Infrastructure.Repositories;
using Serilog;

var builder = Host.CreateApplicationBuilder(args);

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .WriteTo.File(
        @"C:\log\log-.txt",
        rollingInterval: RollingInterval.Day)
    .CreateLogger();

Log.Information("Aplicação iniciada");

builder.Services.AddSingleton<IExcelReader, ExcelReaderService>();

builder.Services.AddSingleton<ImportacaoService>();

builder.Services.AddSingleton<IExcelExporter, ExcelExportService>();

var app = builder.Build();

bool executando = true;

while (executando)
{
    Console.Clear();

    Console.WriteLine("=========================================");
    Console.WriteLine("        R3 INTEGRADOR ARC HOME");
    Console.WriteLine("=========================================");
    Console.WriteLine();
    Console.WriteLine("1 - Processar tabela VAREJO");
    Console.WriteLine("2 - Processar tabela VINILICO");
    Console.WriteLine("3 - Processar tabela LASTRA");
    Console.WriteLine("0 - Sair");
    Console.WriteLine();
    Console.Write("Escolha uma opção: ");

    var opcao = Console.ReadLine();

    switch (opcao)
    {
        case "1":

            await ProcessarAsync(
                app,
                "VAREJO");

            break;

        case "2":

            await ProcessarAsync(
                app,
                "VINILICO");

            break;

        case "3":

            await ProcessarAsync(
                app,
                "LASTRA");

            break;

        case "0":

            executando = false;

            break;

        default:

            Console.WriteLine();
            Console.WriteLine("Opção inválida.");

            break;
    }

    if (executando)
    {
        Console.WriteLine();
        Console.WriteLine("Pressione ENTER para continuar...");

        Console.ReadLine();
    }
}

static async Task ProcessarAsync(
    IHost app,
    string tabela)
{
    Console.WriteLine();
    Console.Write("Informe o caminho do Excel: ");

    var caminho = Console.ReadLine();

    if (string.IsNullOrWhiteSpace(caminho))
    {
        Console.WriteLine("Arquivo inválido.");
        return;
    }

    var service =
        app.Services.GetRequiredService<ImportacaoService>();

    await service.ProcessarAsync(
        caminho,
        tabela);

    Console.WriteLine();
    Console.WriteLine($"Tabela {tabela} processada com sucesso.");
}