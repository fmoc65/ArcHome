using R3Integrador.Application.DTOs;

namespace R3Integrador.Application.Mappers;

public static class ProdutoErpMapper
{
    public static ProdutoErpDto Map(
        ProdutoNormalizado produto)
    {
        return new ProdutoErpDto
        {
            CodigoInterno =
                produto.Referencia,

            CodigoFabrica =
                produto.Referencia,

            DescricaoCompleta =
                $"{produto.Linha} {produto.Formato} {produto.Cor}",

            Grupo =
                produto.TipoTabela,

            SubGrupo =
                produto.Colecao,

            Marca =
                "ARC HOME",

            Unidade =
                "CX",

            PrecoVenda =
                produto.PrecoTabela,

            Ncm =
                "69072100",

            Cst =
                "060",

            Cfop =
                "5102",

            PesoBruto =
                produto.M2Caixa,

            Espessura =
                produto.Espessura
        };
    }
}
