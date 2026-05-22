using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Repositories;

public class ExcelReaderService : IExcelReader
{
    public async Task<List<ProdutoNormalizado>> LerAsync(string caminhoArquivo, string tabela)
    {
        var produtos = new List<ProdutoNormalizado>();
        using var workbook = new XLWorkbook(caminhoArquivo);
        var worksheet = workbook.Worksheet(tabela);

        if (worksheet == null)
        {
            Console.WriteLine($"Aba {tabela} não encontrada.");
            return produtos;
        }

        // Variáveis de estado para manter o valor das células mescladas (Merge)
        string formatoAtual = string.Empty;
        string linhaModeloAtual = string.Empty;

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 0;

        for (int row = 1; row <= ultimaLinha; row++)
        {
            var referencia = worksheet.Cell(row, 2).GetString().Trim();

            // 1. Ignora linhas totalmente vazias
            if (string.IsNullOrWhiteSpace(referencia))
                continue;

            // 2. Filtro inteligente de cabeçalhos: ignora títulos de seções das planilhas
            if (referencia.Equals("Ref.", StringComparison.OrdinalIgnoreCase) || 
                referencia.Equals("VINÍLICOS", StringComparison.OrdinalIgnoreCase) ||
                referencia.Equals("LASTRAS", StringComparison.OrdinalIgnoreCase) ||
                referencia.Equals("Formato", StringComparison.OrdinalIgnoreCase))
                continue;

            // 3. Validação de segurança para códigos de fábrica (Para Varejo/Lastra exige números, para Vinílico aceita Letras)
            if (!tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase))
            {
                if (!referencia.All(char.IsDigit)) 
                    continue;
            }

            // Captura os valores da linha atual do Excel
            var formato = worksheet.Cell(row, 1).GetString().Trim();
            var linhaModelo = worksheet.Cell(row, 3).GetString().Trim();

            // Lógica do Merge para a Coluna 1 (Formato)
            if (!string.IsNullOrWhiteSpace(formato))
            {
                formatoAtual = formato;
            }

            // Lógica do Merge para a Coluna 3 (Linha/Modelo)
            if (!string.IsNullOrWhiteSpace(linhaModelo))
            {
                linhaModeloAtual = linhaModelo;
            }

            // Monta o objeto normalizado utilizando os valores propagados do Merge se necessário
            var produto = new ProdutoNormalizado
            {
                TipoTabela = tabela,
                Formato = formatoAtual,
                Referencia = referencia,
                Linha = linhaModeloAtual, // Aqui garante que herda a "Linha de cima" se estiver mesclado
                Colecao = worksheet.Cell(row, 4).GetString().Trim(),
                Cor = worksheet.Cell(row, 5).GetString().Trim(),
                Superficie = worksheet.Cell(row, 6).GetString().Trim(),
                Faces = ParseInt(worksheet.Cell(row, 7).GetString()),
                VariacaoTonalidade = worksheet.Cell(row, 8).GetString().Trim(),
                M2Caixa = ParseDecimal(worksheet.Cell(row, 11).GetString()),
                Espessura = ParseDecimal(worksheet.Cell(row, 17).GetString())
            };

            // 4. Mapeamento Dinâmico de Preços baseado na Aba
            if (tabela.Equals("VINILICO", StringComparison.OrdinalIgnoreCase))
            {
                produto.PrecoTabela = ParseDecimal(worksheet.Cell(row, 18).GetString());
                produto.PrecoDesconto = ParseDecimal(worksheet.Cell(row, 19).GetString());
                produto.PrecoVenda = produto.PrecoDesconto; // No Vinílico, a última coluna preenchida é a 19
            }
            else
            {
                produto.PrecoTabela = ParseDecimal(worksheet.Cell(row, 18).GetString());
                produto.PrecoDesconto = ParseDecimal(worksheet.Cell(row, 19).GetString());
                produto.PrecoVenda = ParseDecimal(worksheet.Cell(row, 20).GetString()); // Varejo/Lastra buscam o Fracionado na col 20
            }

            produtos.Add(produto);
        }

        Console.WriteLine();
        Console.WriteLine($"[OK] {produtos.Count} produtos processados com sucesso na aba {tabela}.");

        return await Task.FromResult(produtos);
    }

    private static decimal ParseDecimal(string valor)
    {
        if (string.IsNullOrWhiteSpace(valor)) return 0;
        valor = valor.Replace("R$", "").Replace(".", "").Replace(",", ".").Replace("-", "").Trim();
        decimal.TryParse(valor, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var resultado);
        return resultado;
    }

    private static int ParseInt(string valor)
    {
        int.TryParse(valor, out var resultado);
        return resultado;
    }
}