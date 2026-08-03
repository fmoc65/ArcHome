using R3Integrador.Application.Interfaces;
using R3Integrador.Application.Services;
using R3Integrador.Infrastructure.Export;
using R3Integrador.Infrastructure.Repositories;
using R3Integrador.Web.Services;
using Serilog;

var builder = WebApplication.CreateBuilder(args);

Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .WriteTo.Console()
    .CreateLogger();

builder.Host.UseSerilog();
builder.Services.AddRazorComponents().AddInteractiveServerComponents();

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
    ExcelReader = sp.GetRequiredService<IExcelReader>(), VinilicoReader = sp.GetRequiredService<IVinilicoReader>(),
    DelcredereReader = sp.GetRequiredService<IDelcredereReader>(), VillaArtReader = sp.GetRequiredService<IVillaArtReader>(),
    LastraReader = sp.GetRequiredService<ILastraReader>(), RubinettosReader = sp.GetRequiredService<IRubinettosReader>(),
    RocaReader = sp.GetRequiredService<IRocaReader>(), ImersiReader = sp.GetRequiredService<IImersiReader>(),
    StudioMorandinReader = sp.GetRequiredService<IStudioMorandinReader>(), InvitaReader = sp.GetRequiredService<IInvitaReader>(),
    DerossoReader = sp.GetRequiredService<IDerossoReader>(), AtlasReader = sp.GetRequiredService<IAtlasReader>(),
    NinaMartinelliReader = sp.GetRequiredService<INinaMartinelliReader>(), SpecialSlReader = sp.GetRequiredService<ISpecialSlReader>()
});
builder.Services.AddSingleton<ImportacaoService>();
builder.Services.AddSingleton<IExcelExporter, ExcelExportService>();
builder.Services.AddSingleton<RevisaoCamposFiscaisService>();
builder.Services.AddSingleton<HistoricoAtualizacaoStore>();

var app = builder.Build();
app.UseExceptionHandler("/Error", createScopeForErrors: true);
app.UseStaticFiles();
app.UseAntiforgery();
app.MapRazorComponents<R3Integrador.Web.Components.App>().AddInteractiveServerRenderMode();
app.Run();
