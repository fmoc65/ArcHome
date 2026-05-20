using R3Integrador.Application.DTOs;
using R3Integrador.Core.Enums;

namespace R3Integrador.Application.Mappers;

public static class ProdutoErpMapper
{
    public static ProdutoErpDto Map(ProdutoNormalizado produto)
    {
        var grupo = IdentificarGrupo(produto.TipoTabela);
        var subGrupo = IdentificarSubGrupo(produto.Colecao, grupo);

        string apenasNumeros = new string(produto.Referencia.Where(char.IsDigit).ToArray());
        string sufixoCodigo = apenasNumeros.Length >= 3 
            ? apenasNumeros[^3..] 
            : apenasNumeros.PadLeft(3, '0');
        
        string codigoInternoMascarado = $"{(int)grupo}.{(int)subGrupo}.{sufixoCodigo}";
        string descricaoGerada = $"{produto.Linha} {produto.Colecao} {produto.Superficie} {produto.Cor} {produto.Formato}".Trim().ToUpper();

        return new ProdutoErpDto
        {
            CodigoInterno = codigoInternoMascarado,
            CodigoFabrica = produto.Referencia,
            CodigoBarras = string.Empty,
            DescricaoCompleta = descricaoGerada,
            DescricaoComercial = $"{produto.Linha} {produto.Colecao}".ToUpper(),
            Grupo = ((int)grupo).ToString(),
            SubGrupo = ((int)subGrupo).ToString(),
            Marca = "ARC HOME",
            Linha = produto.Linha,
            Modelo = produto.Colecao,
            Voltagem = "N/A",
            Cor = produto.Cor,
            Ncm = "69072100",
            UfOrigem = "SP",
            
            // Mapeamento correto dos preços nas colunas oficiais do ERP
            PrecoVenda = produto.PrecoVenda, 
            PrecoFabrica = produto.PrecoTabela,
            DescontoPercentual = 0,
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

    private static Grupo IdentificarGrupo(string tipoTabela)
    {
        if (string.IsNullOrWhiteSpace(tipoTabela)) return Grupo.NaoDefinido;
        string t = tipoTabela.ToUpper();
        if (t.Contains("VAREJO") || t.Contains("DELCREDERE")) return Grupo.Revestimento;
        if (t.Contains("VINILICO")) return Grupo.Vinilico;
        if (t.Contains("LASTRA")) return Grupo.Lastra;
        if (t.Contains("BOUTIQUE") || t.Contains("VILLA")) return Grupo.Boutique;
        
        return Grupo.NaoDefinido;
    }

    private static SubGrupo IdentificarSubGrupo(string colecao, Grupo grupo)
    {
        if (string.IsNullOrWhiteSpace(colecao)) return SubGrupo.Padrao;
        if (grupo == Grupo.Vinilico)
        {
            if (colecao.Contains("SPC", StringComparison.OrdinalIgnoreCase)) return SubGrupo.Spc;
            if (colecao.Contains("LVT", StringComparison.OrdinalIgnoreCase)) return SubGrupo.Lvt;
        }
        if (colecao.Contains("PORCELANATO", StringComparison.OrdinalIgnoreCase)) return SubGrupo.Porcelanato;
        
        return SubGrupo.Padrao;
    }
}