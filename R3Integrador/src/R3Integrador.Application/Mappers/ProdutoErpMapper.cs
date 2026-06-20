using R3Integrador.Application.DTOs;
using System.Text.RegularExpressions;

namespace R3Integrador.Application.Mappers;

public static class ProdutoErpMapper
{
    public static ProdutoErpDto Map(ProdutoNormalizado produto)
    {
        var descricaoGerada = GerarDescricao(produto);

        return new ProdutoErpDto
        {
            CodigoInterno = string.Empty,
            CodigoFabrica = produto.Referencia,
            CodigoBarras = string.Empty,
            DescricaoCompleta = descricaoGerada,
            DescricaoComercial = GerarDescricaoComercial(produto),
            Grupo = produto.Grupo,
            SubGrupo = produto.SubGrupo,
            Marca = produto.Marca,
            Linha = produto.Linha,
            Modelo = produto.Modelo,
            Voltagem = string.Empty,
            Cor = produto.Cor,
            Ncm = ObterNcm(produto),
            UfOrigem = "SP",
            
            // Mapeamento correto dos preÃ§os nas colunas oficiais do ERP
            PrecoVenda = produto.PrecoVenda, 
            PrecoFabrica = produto.PrecoTabela,
            DescontoPercentual = CalcularDescontoPercentual(produto),
            IpiPercentual = 0,
            AliqIcmsOrigem = 18.00m,
            AliqIcmsInterna = 18.00m,
            Iva = 0,
            FreteReais = 0,
            FretePercentual = 0,
            Unidade = "M2",
            QtdeEmbalagemVenda = ObterM2PorCaixa(produto),
            
            Cst = "060",
            AliquotaCofinsCst = "01",
            AliquotaIpiCst = "49",
            AliquotaPisCst = "01",
            Csosn = "500",
            CfopDentro = "5405",
            CfopFora = "6404",
            
            PesoLiquido = 0,
            PesoBruto = 0,
            QtdeEmbalagemCompra = 1,
            ValorPi = 0,
            AliquotaCofins = 0,
            AliquotaPis = 0,
            PercentualSt = 0,
            UnidFabril = "CX",
            
            // Dica: Como o ERP nÃ£o tem coluna de "Espessura", guardamos essa informaÃ§Ã£o na ObservaÃ§Ã£o!
            Observacao = $"Importado via R3Integrador - Tabela {produto.TipoTabela} - Espessura: {produto.Espessura}mm",
            
            DiferencaIcms = 0,
            ReducaoBaseIcms = 0,
            ReducaoBaseSt = 0,
            EnquadramentoIpi = ObterEnquadramentoIpi(produto),
            AliquotaIbs = ObterAliquotaIbs(produto),
            AliquotaCbs = ObterAliquotaCbs(produto),
            ClassificacaoTributaria = ObterClassificacaoTributaria(produto)
            
            // <-- AS LINHAS "PrecoFracionado = ..." E "Espessura = ..." FORAM REMOVIDAS DAQUI
            // Pois elas nÃ£o existem no DTO de 60 colunas exigido pelo layout do ERP.
        };
    }

    private static string ObterNcm(ProdutoNormalizado produto)
    {
        if (!EhPorcelanato(produto))
        {
            return "0";
        }

        return produto.Referencia.Equals("120003", StringComparison.OrdinalIgnoreCase)
            ? "69072200"
            : "69072100";
    }

    private static string ObterEnquadramentoIpi(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) ? "999" : "0";
    }

    private static string ObterAliquotaIbs(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) ? "0,1" : "0";
    }

    private static string ObterAliquotaCbs(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) ? "0,9" : "0";
    }

    private static string ObterClassificacaoTributaria(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) ? "000001" : "0";
    }

    private static bool EhPorcelanato(ProdutoNormalizado produto)
    {
        return produto.Grupo.Equals("PORCELANATO", StringComparison.OrdinalIgnoreCase);
    }

    private static decimal ObterM2PorCaixa(ProdutoNormalizado produto)
    {
        return produto.M2Caixa > 0 ? produto.M2Caixa : 1;
    }

    private static string GerarDescricao(ProdutoNormalizado produto)
    {
        var descricao = $"{produto.Grupo} {produto.Linha} {produto.SubGrupo} {produto.Cor} {produto.Modelo}";
        return NormalizarEspacos(descricao).ToUpper();
    }

    private static string GerarDescricaoComercial(ProdutoNormalizado produto)
    {
        var descricao = $"{produto.Grupo} {produto.Modelo} {produto.Cor}";
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

    private static string NormalizarEspacos(string valor)
    {
        return Regex.Replace(valor.Trim(), @"\s+", " ");
    }

}



