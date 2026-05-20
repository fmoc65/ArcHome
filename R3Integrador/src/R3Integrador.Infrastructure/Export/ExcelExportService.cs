using ClosedXML.Excel;
using R3Integrador.Application.DTOs;
using R3Integrador.Application.Interfaces;

namespace R3Integrador.Infrastructure.Export;

public class ExcelExportService : IExcelExporter
{
    public async Task ExportarAsync(List<ProdutoErpDto> produtos, string caminhoSaida)
    {
        // Força a criação das pastas caso o Windows bloqueie o caminho
        var diretorio = Path.GetDirectoryName(caminhoSaida);
        if (!string.IsNullOrEmpty(diretorio) && !Directory.Exists(diretorio))
        {
            Directory.CreateDirectory(diretorio);
        }

        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("IMPORTACAO_ERP");

        CriarCabecalho(worksheet);

        int linha = 2;

        foreach (var produto in produtos)
        {
            // Escrevendo as colunas exatamente na ordem posicional do Layout de 60 colunas do ERP
            worksheet.Cell(linha, 1).Value = produto.CodigoInterno;
            worksheet.Cell(linha, 2).Value = produto.CodigoFabrica;
            worksheet.Cell(linha, 3).Value = produto.CodigoBarras;
            worksheet.Cell(linha, 4).Value = produto.DescricaoCompleta;
            worksheet.Cell(linha, 5).Value = produto.DescricaoComercial;
            worksheet.Cell(linha, 6).Value = produto.Grupo;
            worksheet.Cell(linha, 7).Value = produto.SubGrupo;
            worksheet.Cell(linha, 8).Value = produto.Marca;
            worksheet.Cell(linha, 9).Value = produto.Linha;
            worksheet.Cell(linha, 10).Value = produto.Modelo;
            worksheet.Cell(linha, 11).Value = produto.Voltagem;
            worksheet.Cell(linha, 12).Value = produto.Cor;
            worksheet.Cell(linha, 13).Value = produto.Ncm;
            worksheet.Cell(linha, 14).Value = produto.UfOrigem;
            worksheet.Cell(linha, 15).Value = produto.PrecoVenda;
            worksheet.Cell(linha, 16).Value = produto.PrecoFabrica;
            worksheet.Cell(linha, 17).Value = produto.DescontoPercentual;
            worksheet.Cell(linha, 18).Value = produto.IpiPercentual;
            worksheet.Cell(linha, 19).Value = produto.AliqIcmsOrigem;
            worksheet.Cell(linha, 20).Value = produto.AliqIcmsInterna;
            worksheet.Cell(linha, 21).Value = produto.Iva;
            worksheet.Cell(linha, 22).Value = produto.FreteReais;
            worksheet.Cell(linha, 23).Value = produto.FretePercentual;
            worksheet.Cell(linha, 24).Value = produto.Unidade;
            worksheet.Cell(linha, 25).Value = produto.QtdeEmbalagemVenda;
            worksheet.Cell(linha, 26).Value = produto.Cst;
            worksheet.Cell(linha, 27).Value = produto.AliquotaCofinsCst;
            worksheet.Cell(linha, 28).Value = produto.AliquotaIpiCst;
            worksheet.Cell(linha, 29).Value = produto.AliquotaPisCst;
            worksheet.Cell(linha, 30).Value = produto.Csosn;
            worksheet.Cell(linha, 31).Value = produto.CfopDentro; // Coluna oficial de CFOP interna
            worksheet.Cell(linha, 32).Value = produto.CfopFora;   // Coluna oficial de CFOP externa
            worksheet.Cell(linha, 33).Value = produto.PesoLiquido;
            worksheet.Cell(linha, 34).Value = produto.PesoBruto;
            worksheet.Cell(linha, 35).Value = produto.QtdeEmbalagemCompra;
            worksheet.Cell(linha, 36).Value = produto.ValorPi;
            worksheet.Cell(linha, 37).Value = produto.AliquotaCofins;
            worksheet.Cell(linha, 38).Value = produto.AliquotaPis;
            worksheet.Cell(linha, 39).Value = produto.PercentualSt;
            worksheet.Cell(linha, 40).Value = produto.UnidFabril;
            worksheet.Cell(linha, 41).Value = produto.Observacao;
            worksheet.Cell(linha, 42).Value = produto.DiferencaIcms;
            worksheet.Cell(linha, 43).Value = produto.ReducaoBaseIcms;
            worksheet.Cell(linha, 44).Value = produto.ReducaoBaseSt;
            worksheet.Cell(linha, 45).Value = produto.RetencaoPis;
            worksheet.Cell(linha, 46).Value = produto.RetencaoCofins;
            worksheet.Cell(linha, 47).Value = produto.RetencaoCsll;
            worksheet.Cell(linha, 48).Value = produto.RetencaoIrrf;
            worksheet.Cell(linha, 49).Value = produto.RetencaoPrevSocial;
            worksheet.Cell(linha, 50).Value = produto.Localizacao;
            worksheet.Cell(linha, 51).Value = produto.EnquadramentoIpi;
            worksheet.Cell(linha, 52).Value = produto.AliquotaPisOrigem;
            worksheet.Cell(linha, 53).Value = produto.AliquotaCofinsOrigem;
            worksheet.Cell(linha, 54).Value = produto.Imagem;
            worksheet.Cell(linha, 55).Value = produto.EstoqueMinimo;
            worksheet.Cell(linha, 56).Value = produto.EstoqueMaximo;
            worksheet.Cell(linha, 57).Value = produto.AliquotaIbs;
            worksheet.Cell(linha, 58).Value = produto.AliquotaCbs;
            worksheet.Cell(linha, 59).Value = produto.ClassificacaoTributaria;
            worksheet.Cell(linha, 60).Value = produto.CodigoBeneficio;

            linha++;
        }

        // ... (código anterior do preenchimento das células)

        worksheet.Columns().AdjustToContents();

        // CORREÇÃO AQUI: Força a liberação física e imediata do arquivo para o Windows
        using (var fileStream = new FileStream(caminhoSaida, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            workbook.SaveAs(fileStream);
            fileStream.Flush(true); // <--- Força o Windows a gravar fisicamente no HD/SSD AGORA
        }

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"[Sucesso] Arquivo de importação gerado em: {caminhoSaida}");
        Console.ResetColor();

        // Removido o Task.CompletedTask inútil que causava o falso positivo antes do término real
    }

    private static void CriarCabecalho(IXLWorksheet worksheet)
    {
        string[] cabecalhos = new string[]
        {
            "CÓDIGO INTERNO", "CÓDIGO FÁBRICA", "CODIGOBARRAS", "DESCRIÇÃO COMPLETA", "DESCRIÇÃO COMERCIAL",
            "GRUPO", "SUBGRUPO", "MARCA", "LINHA", "MODELO", "VOLTAGEM", "COR", "NCM", "UF ORIGEM",
            "PREÇO VENDA", "PREÇO DE FÁBRICA", "DESCONTO %", "IPI %", "ALIQICMSORIGEM", "ALIQICMSINTERNA",
            "IVA", "FRETE R$", "FRETE %", "UNIDADE", "QTDE EMBALAGEM DE VENDA", "CST", "ALIQUOTA COFINS CST",
            "ALIQUOTA IPI CST", "ALIQUOTA PIS CST", "CSOSN", "CFOP DENTRO", "CFOP FORA", "PESOLIQ", "PESOBRUTO",
            "QTDE EMBALAGEM DE COMPRA", "VALOR PI", "ALIQUOTA COFINS", "ALIQUOTA PIS", "PERCENTUAL ST",
            "UNID FABRIL", "OBSERVACAO", "DIFERENÇA ICMS", "REDUÇÃO BASE ICMS", "REDUÇÃO BASE ST",
            "RETENÇÃO PIS", "RETENÇÃO COFINS", "RETENÇÃO CSLL", "RETENÇÃO IRRF", "RETENÇÃO PREV. SOCIAL",
            "LOCALIZAÇÃO", "ENQUADRAMENTO IPI", "ALIQUOTA PIS ORIGEM", "ALIQUOTA COFINS ORIGEM", "IMAGEM",
            "ESTOQUE MINIMO", "ESTOQUE MAXIMO", "ALIQUOTA IBS", "ALIQUOTA CBS", "CLASSIFICACAO TRIBUTARIA",
            "CODIGO BENEFICIO"
        };

        for (int i = 0; i < cabecalhos.Length; i++)
        {
            var cell = worksheet.Cell(1, i + 1);
            cell.Value = cabecalhos[i];
            cell.Style.Font.Bold = true;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#1F4E78");
            cell.Style.Font.SetFontColor(XLColor.White);
        }
    }
}