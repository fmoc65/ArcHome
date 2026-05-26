using R3Integrador.Application.DTOs;
using System.Text.RegularExpressions;

namespace R3Integrador.Application.Mappers;

public static class ProdutoErpMapper
{
    public static ProdutoErpDto Map(ProdutoNormalizado produto)
    {
        var grupo = GerarGrupo(produto);
        var subGrupo = GerarSubGrupo(produto);
        var descricaoGerada = GerarDescricao(produto);

        return new ProdutoErpDto
        {
            CodigoInterno = string.Empty,
            CodigoFabrica = produto.Referencia,
            CodigoBarras = string.Empty,
            DescricaoCompleta = descricaoGerada,
            DescricaoComercial = GerarDescricaoComercial(produto, grupo),
            Grupo = grupo,
            SubGrupo = subGrupo,
            Marca = "ARC HOME",
            Linha = produto.Linha,
            Modelo = GerarModelo(produto),
            Voltagem = "N/A",
            Cor = produto.Cor,
            Ncm = "69072100",
            UfOrigem = "SP",
            
            // Mapeamento correto dos preços nas colunas oficiais do ERP
            PrecoVenda = produto.PrecoVenda, 
            PrecoFabrica = produto.PrecoTabela,
            DescontoPercentual = CalcularDescontoPercentual(produto),
            IpiPercentual = 0,
            AliqIcmsOrigem = 18.00m,
            AliqIcmsInterna = 18.00m,
            Iva = 0,
            FreteReais = 0,
            FretePercentual = 0,
            Unidade = "CX",
            QtdeEmbalagemVenda = 1,
            
            Cst = "060",
            AliquotaCofinsCst = "01",
            AliquotaIpiCst = "49",
            AliquotaPisCst = "01",
            Csosn = "500",
            CfopDentro = "5405",
            CfopFora = "6404",
            
            PesoLiquido = 0,
            PesoBruto = produto.M2Caixa, // Se o ERP exigir o peso por caixa, mapeamos aqui
            QtdeEmbalagemCompra = 1,
            ValorPi = 0,
            AliquotaCofins = 0,
            AliquotaPis = 0,
            PercentualSt = 0,
            UnidFabril = "CX",
            
            // Dica: Como o ERP não tem coluna de "Espessura", guardamos essa informação na Observação!
            Observacao = $"Importado via R3Integrador - Tabela {produto.TipoTabela} - Espessura: {produto.Espessura}mm",
            
            DiferencaIcms = 0,
            ReducaoBaseIcms = 0,
            ReducaoBaseSt = 0
            
            // <-- AS LINHAS "PrecoFracionado = ..." E "Espessura = ..." FORAM REMOVIDAS DAQUI
            // Pois elas não existem no DTO de 60 colunas exigido pelo layout do ERP.
        };
    }

    private static string GerarGrupo(ProdutoNormalizado produto)
    {
        return NormalizarEspacos(produto.TipoTabela).ToUpper();
    }

    private static string GerarSubGrupo(ProdutoNormalizado produto)
    {
        return NormalizarEspacos($"{produto.Colecao} {produto.Superficie}").ToUpper();
    }

    private static string GerarModelo(ProdutoNormalizado produto)
    {
        return NormalizarEspacos($"{produto.Colecao} {produto.TabelaPreco}").ToUpper();
    }

    private static string GerarDescricao(ProdutoNormalizado produto)
    {
        var tipoProduto = produto.TipoTabela.Contains("VINILICO", StringComparison.OrdinalIgnoreCase)
            ? "VINILICO"
            : "PORCELANATO";

        var descricao = $"{tipoProduto} {produto.Linha} {produto.Superficie} {produto.Cor} {NormalizarMedida(produto.Formato)}";
        return NormalizarEspacos(descricao).ToUpper();
    }

    private static string GerarDescricaoComercial(ProdutoNormalizado produto, string grupo)
    {
        var descricao = $"{grupo} {NormalizarMedida(produto.Formato)} {produto.Cor}";
        return NormalizarEspacos(descricao).ToUpper();
    }

    private static decimal CalcularDescontoPercentual(ProdutoNormalizado produto)
    {
        if (produto.PrecoTabela <= 0 || produto.PrecoDesconto <= 0)
        {
            return 0;
        }

        var desconto = (produto.PrecoTabela - produto.PrecoDesconto) / produto.PrecoTabela * 100;
        return Math.Round(desconto, 2);
    }

    private static string NormalizarMedida(string formato)
    {
        return formato.Replace(" ", string.Empty).ToUpper();
    }

    private static string NormalizarEspacos(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ");
    }

}
