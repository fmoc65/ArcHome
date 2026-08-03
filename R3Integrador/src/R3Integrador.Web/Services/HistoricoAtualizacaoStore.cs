namespace R3Integrador.Web.Services;

public sealed class HistoricoAtualizacaoStore
{
    private readonly List<AtualizacaoTabela> itens = [];
    public IReadOnlyList<AtualizacaoTabela> Itens => itens.OrderByDescending(item => item.ProcessadoEm).ToList();
    public void Registrar(string tabela, string arquivo) => itens.Add(new AtualizacaoTabela(tabela, arquivo, DateTime.Now));
}

public sealed record AtualizacaoTabela(string Tabela, string Arquivo, DateTime ProcessadoEm);
