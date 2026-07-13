using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace R3Integrador.Infrastructure.Repositories;

// O PDF nao contem uma camada tabular. O pdftohtml preserva as coordenadas de cada
// texto, permitindo reconstruir codigo, descricao, formato, linha e preco publicado.
public sealed class SpecialSlReaderService : ISpecialSlReader
{
    private const string Marca = "SL";
    private static readonly CultureInfo CulturaPtBr = CultureInfo.GetCultureInfo("pt-BR");

    public async Task<List<ProdutoErpDto>> LerAsync(string caminhoArquivo)
    {
        var caminhoXmlBase = Path.Combine(Path.GetTempPath(), $"r3-sl-{Guid.NewGuid():N}");
        var caminhoXml = $"{caminhoXmlBase}.xml";

        try
        {
            await GerarXmlDoPdfAsync(caminhoArquivo, caminhoXmlBase);
            var produtos = ExtrairProdutos(XDocument.Load(caminhoXml));

            Console.WriteLine();
            Console.WriteLine($"[OK] {produtos.Count} produtos da Tabela Especial SL processados.");
            Console.WriteLine("[ATENCAO] O PDF nao informa NCM, EAN ou tributacao por item; arquivo permanece provisorio.");
            return produtos;
        }
        finally
        {
            ApagarArquivosTemporarios(caminhoXmlBase);
        }
    }

    private static async Task GerarXmlDoPdfAsync(string caminhoArquivo, string caminhoXmlBase)
    {
        var inicio = new ProcessStartInfo("pdftohtml")
        {
            RedirectStandardError = true,
            RedirectStandardOutput = true,
            UseShellExecute = false
        };
        inicio.ArgumentList.Add("-xml");
        inicio.ArgumentList.Add("-hidden");
        inicio.ArgumentList.Add("-enc");
        inicio.ArgumentList.Add("UTF-8");
        inicio.ArgumentList.Add(caminhoArquivo);
        inicio.ArgumentList.Add(caminhoXmlBase);

        try
        {
            using var processo = Process.Start(inicio)
                ?? throw new InvalidOperationException("Nao foi possivel iniciar o conversor do PDF.");
            var erro = processo.StandardError.ReadToEndAsync();
            await processo.WaitForExitAsync();

            if (processo.ExitCode != 0)
            {
                throw new InvalidOperationException($"Falha ao ler PDF: {await erro}");
            }
        }
        catch (Win32Exception ex)
        {
            throw new InvalidOperationException(
                "O leitor da Tabela Especial SL requer o utilitario 'pdftohtml' (pacote poppler-utils) instalado no sistema.", ex);
        }
    }

    private static List<ProdutoErpDto> ExtrairProdutos(XDocument documento)
    {
        var produtos = new List<ProdutoErpDto>();
        var referenciasUsadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var pagina in documento.Root!.Elements("page"))
        {
            var textos = pagina.Elements("text")
                .Select(elemento => new TextoPdf(
                    (int)elemento.Attribute("top")!,
                    (int)elemento.Attribute("left")!,
                    (int)elemento.Attribute("font")!,
                    NormalizarTexto(elemento.Value)))
                .Where(texto => !string.IsNullOrWhiteSpace(texto.Valor))
                .ToList();

            var precos = textos.Where(texto => texto.Font == 3 && LerPreco(texto.Valor) > 0).ToList();
            var codigos = textos.Where(texto => Regex.IsMatch(texto.Valor, @"^(CT\s*)?\d{4,5}$"))
                .OrderBy(texto => texto.Top)
                .ThenBy(texto => texto.Left);

            foreach (var codigo in codigos)
            {
                var colunaDireita = codigo.Left >= 450;
                var textosDaColuna = textos.Where(texto => (texto.Left >= 450) == colunaDireita).ToList();
                var descricao = textosDaColuna
                    .Where(texto => Math.Abs(texto.Top - codigo.Top) <= 3 && texto.Left > codigo.Left + 35 && texto.Font != 3)
                    .OrderBy(texto => texto.Left)
                    .FirstOrDefault()?.Valor;

                if (string.IsNullOrWhiteSpace(descricao) || descricao is "CODIGO" or "DESCRICAO")
                {
                    continue;
                }

                var formato = textosDaColuna
                    .Where(texto => texto.Font == 0 && texto.Top <= codigo.Top && Regex.IsMatch(texto.Valor, @"^\d+\s*X\s*\d+$"))
                    .OrderByDescending(texto => texto.Top)
                    .FirstOrDefault()?.Valor ?? string.Empty;
                var linha = textosDaColuna
                    .Where(texto => texto.Font == 2 && texto.Top <= codigo.Top && texto.Left < codigo.Left)
                    .OrderByDescending(texto => texto.Top)
                    .FirstOrDefault()?.Valor ?? string.Empty;
                var preco = precos
                    .Where(texto => (texto.Left >= 450) == colunaDireita)
                    .OrderBy(texto => Math.Abs(texto.Top - codigo.Top))
                    .Select(texto => LerPreco(texto.Valor))
                    .FirstOrDefault();

                if (preco <= 0)
                {
                    continue;
                }

                var referencia = CriarReferenciaUnica(codigo.Valor, formato, preco, referenciasUsadas);
                var unidade = linha.Equals("CANTONEIRAS", StringComparison.OrdinalIgnoreCase) ? "PC" : "M2";

                produtos.Add(new ProdutoErpDto
                {
                    CodigoFabrica = referencia,
                    DescricaoCompleta = CriarDescricao(descricao, linha, formato),
                    DescricaoComercial = descricao,
                    Grupo = "REVESTIMENTOS",
                    SubGrupo = string.IsNullOrWhiteSpace(linha) ? "REVESTIMENTOS" : linha,
                    Marca = Marca,
                    Linha = linha,
                    Modelo = formato,
                    UfOrigem = string.Empty,
                    // A tabela declara preco FOB. Sem frete/margem homologados, o valor e
                    // apenas preservado nas duas colunas para revisao comercial posterior.
                    PrecoVenda = preco,
                    PrecoFabrica = preco,
                    Unidade = unidade,
                    QtdeEmbalagemVenda = 1,
                    QtdeEmbalagemCompra = 1,
                    UnidFabril = unidade,
                    EstoqueMinimo = 0,
                    EstoqueMaximo = 0,
                    Observacao = $"Extraido do PDF Tabela 2026 Especial SL - preco FOB Tatuí - sem IPI/ST/DIFAL declarados na origem - formato {formato} - embalagem, peso, EAN, NCM, ICMS, CST, CFOP e regra comercial pendentes de homologacao",
                    SituacaoCamposFiscais = CriarSituacaoCamposFiscais()
                });
            }
        }

        return produtos;
    }

    private static Dictionary<int, SituacaoCampoFiscal> CriarSituacaoCamposFiscais() => new()
    {
        [13] = SituacaoCampoFiscal.Pendente, [18] = SituacaoCampoFiscal.Pendente,
        [19] = SituacaoCampoFiscal.Pendente, [20] = SituacaoCampoFiscal.Pendente,
        [21] = SituacaoCampoFiscal.Pendente, [26] = SituacaoCampoFiscal.Pendente,
        [27] = SituacaoCampoFiscal.Pendente, [28] = SituacaoCampoFiscal.Pendente,
        [29] = SituacaoCampoFiscal.Pendente, [30] = SituacaoCampoFiscal.Pendente,
        [31] = SituacaoCampoFiscal.Pendente, [32] = SituacaoCampoFiscal.Pendente,
        [39] = SituacaoCampoFiscal.Pendente, [51] = SituacaoCampoFiscal.Pendente,
        [52] = SituacaoCampoFiscal.Pendente, [53] = SituacaoCampoFiscal.Pendente,
        [57] = SituacaoCampoFiscal.Pendente, [58] = SituacaoCampoFiscal.Pendente,
        [59] = SituacaoCampoFiscal.Pendente, [60] = SituacaoCampoFiscal.Pendente
    };

    private static string CriarReferenciaUnica(string codigo, string formato, decimal preco, ISet<string> usadas)
    {
        var referencia = codigo.Replace(" ", string.Empty);
        if (usadas.Add(referencia)) return referencia;

        var baseReferencia = $"{referencia}-{Regex.Replace(formato, @"\D", string.Empty)}-{preco:0.##}".Replace(',', '-');
        referencia = baseReferencia;
        var sequencia = 2;
        while (!usadas.Add(referencia)) referencia = $"{baseReferencia}-{sequencia++}";
        return referencia;
    }

    private static decimal LerPreco(string valor) => decimal.TryParse(valor, NumberStyles.Number, CulturaPtBr, out var preco) ? preco : 0;
    private static string CriarDescricao(params string[] valores) => string.Join(" - ", valores.Where(valor => !string.IsNullOrWhiteSpace(valor)));
    private static string NormalizarTexto(string valor) => Regex.Replace(valor, @"\s+", " ").Trim().ToUpperInvariant();

    private static void ApagarArquivosTemporarios(string caminhoBase)
    {
        foreach (var caminho in Directory.EnumerateFiles(Path.GetTempPath(), $"{Path.GetFileName(caminhoBase)}*"))
        {
            File.Delete(caminho);
        }
    }

    private sealed record TextoPdf(int Top, int Left, int Font, string Valor);
}
