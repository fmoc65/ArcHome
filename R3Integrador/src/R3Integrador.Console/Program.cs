using R3Integrador.Application.Services;
using Serilog;

var builder = Host.CreateApplicationBuilder(args);

Log.Logger = new LoggerConfiguration()
    .WriteTo.Console()
    .WriteTo.File(
        "Logs/log-.txt",
        rollingInterval: RollingInterval.Day)
    .CreateLogger();

builder.Services.AddSingleton<ImportacaoService>();

var app = builder.Build();

Console.WriteLine("========================================");
Console.WriteLine(" INTEGRADOR ARC HOME ");
Console.WriteLine("========================================");
Console.WriteLine();
Console.WriteLine("1 - Processar tabela fornecedor");
Console.WriteLine("0 - Sair");
Console.WriteLine();
Console.Write("Escolha: ");

var opcao = Console.ReadLine();

switch (opcao)
{
    case "1":

        Console.Write("Informe o caminho do Excel: ");

        var caminho = Console.ReadLine();

        var service = app.Services.GetRequiredService<ImportacaoService>();

        await service.ProcessarAsync(caminho!);

        break;

    case "0":
        return;
}