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
            
            PrecoVenda = produto.PrecoVenda, 
            PrecoFabrica = produto.PrecoTabela,
            DescontoPercentual = CalcularDescontoPercentual(produto),
            IpiPercentual = produto.IpiPercentual ?? ObterIpiPercentual(produto),
            AliqIcmsOrigem = 12.00m,
            AliqIcmsInterna = ObterAliqIcmsInterna(produto),
            Iva = ObterIva(produto),
            FreteReais = 0,
            FretePercentual = 0,
            Unidade = "M2",
            QtdeEmbalagemVenda = 0,
            
            Cst = ObterCst(produto),
            AliquotaCofinsCst = "01",
            AliquotaIpiCst = "49",
            AliquotaPisCst = "01",
            Csosn = "500",
            CfopDentro = "5405",
            CfopFora = "6404",
            
            PesoLiquido = 0,
            PesoBruto = produto.PesoBrutoM2,
            QtdeEmbalagemCompra = 1,
            ValorPi = 0,
            AliquotaCofins = 0,
            AliquotaPis = 0,
            PercentualSt = ObterPercentualSt(produto),
            UnidFabril = "CX",
            
            Observacao = $"Importado via R3Integrador - Tabela {produto.TipoTabela} - Espessura: {produto.Espessura}mm",
            
            DiferencaIcms = 0,
            ReducaoBaseIcms = 0,
            ReducaoBaseSt = 0,
            EnquadramentoIpi = !string.IsNullOrWhiteSpace(produto.EnquadramentoIpi) ? produto.EnquadramentoIpi : ObterEnquadramentoIpi(produto),
            AliquotaIbs = !string.IsNullOrWhiteSpace(produto.AliquotaIbs) ? produto.AliquotaIbs : ObterAliquotaIbs(produto),
            AliquotaCbs = !string.IsNullOrWhiteSpace(produto.AliquotaCbs) ? produto.AliquotaCbs : ObterAliquotaCbs(produto),
            ClassificacaoTributaria = !string.IsNullOrWhiteSpace(produto.ClassificacaoTributaria) ? produto.ClassificacaoTributaria : ObterClassificacaoTributaria(produto),
            CodigoBeneficio = produto.CodigoBeneficio,
            AliquotaPisOrigem = ObterAliquotaPisOrigem(produto),
            AliquotaCofinsOrigem = ObterAliquotaCofinsOrigem(produto)
        };
    }

    private static decimal ObterAliqIcmsInterna(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) || EhVinilico(produto) ? 12.00m : 18.00m;
    }

    private static decimal ObterIva(ProdutoNormalizado produto)
    {
        if (EhVinilico(produto))
        {
            return 66.00m;
        }

        return EhPorcelanato(produto) ? 81.00m : 0;
    }

    private static decimal ObterPercentualSt(ProdutoNormalizado produto)
    {
        if (EhVinilico(produto))
        {
            return 7.92m;
        }

        return EhPorcelanato(produto) ? 9.86m : 0;
    }

    private static string ObterAliquotaPisOrigem(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) || EhVinilico(produto) ? "0,65" : string.Empty;
    }

    private static string ObterAliquotaCofinsOrigem(ProdutoNormalizado produto)
    {
        return EhPorcelanato(produto) || EhVinilico(produto) ? "3" : string.Empty;
    }

    private static string ObterNcm(ProdutoNormalizado produto)
    {
        if (EhVinilico(produto))
        {
            return "39181000";
        }

        if (!EhPorcelanato(produto))
        {
            return "0";
        }

        return produto.Referencia.Equals("120003", StringComparison.OrdinalIgnoreCase)
            ? "69072200"
            : "69072100";
    }

    private static decimal ObterIpiPercentual(ProdutoNormalizado produto)
    {
        if (!EhPorcelanato(produto))
        {
            return 0;
        }

        var ncm = ObterNcm(produto);
        return ncm.Equals("69072100", StringComparison.OrdinalIgnoreCase) ? 0.65m : 0;
    }

    private static string ObterCst(ProdutoNormalizado produto)
    {
        if (EhPorcelanato(produto) || EhVinilico(produto))
        {
            return "010";
        }

        return "060";
    }

    private static bool EhVinilico(ProdutoNormalizado produto)
    {
        return produto.Grupo.Equals("VINILICO", StringComparison.OrdinalIgnoreCase);
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

