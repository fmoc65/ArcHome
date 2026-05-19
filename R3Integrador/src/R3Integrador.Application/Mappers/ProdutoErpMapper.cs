using R3Integrador.Application.DTOs;
using R3Integrador.Core.Enums;

namespace R3Integrador.Application.Mappers;

public static class ProdutoErpMapper
{
    public static ProdutoErpDto Map(ProdutoNormalizado produto)
    {
        // 1. Identifica Grupo e Subgrupo por Enums baseados nas regras da ARC HOME
        var grupo = IdentificarGrupo(produto.TipoTabela);
        var subGrupo = IdentificarSubGrupo(produto.Colecao, grupo);

        // 2. Extrai apenas números da referência original (ex: "SPC1850001" vira "1850001")
        string apenasNumeros = new string(produto.Referencia.Where(char.IsDigit).ToArray());
        
        // Pega os 3 últimos dígitos para evitar repetição direta do código longo da fábrica
        string sufixoCodigo = apenasNumeros.Length >= 3 
            ? apenasNumeros[^3..] 
            : apenasNumeros.PadLeft(3, '0');

        // 3. Monta a máscara inteligente interna exigida: XX.XX.XXX (Ex: 10.11.001)
        // Usamos o (int) para pegar o número do Enum na máscara
        string codigoInternoMascarado = $"{(int)grupo}.{(int)subGrupo}.{sufixoCodigo}";

        return new ProdutoErpDto
        {
            CodigoInterno = codigoInternoMascarado, 
            CodigoFabrica = produto.Referencia, // Mantém a referência original intacta para conferência
            DescricaoCompleta = $"{produto.Linha} {produto.Formato} {produto.Cor}".Trim().ToUpper(),
            DescricaoComercial = $"{produto.Linha} {produto.Cor}".Trim().ToUpper(),
            
            // Ajustado para gravar o NÚMERO do grupo/subgrupo como string, ideal para o banco do ERP
            Grupo = ((int)grupo).ToString(),
            SubGrupo = ((int)subGrupo).ToString(),
            
            Marca = "ARC HOME",
            Linha = produto.Linha ?? string.Empty,
            Cor = produto.Cor ?? string.Empty,
            Unidade = "CX",
            
            // Precos
            PrecoVenda = produto.PrecoTabela,
            PrecoFabrica = produto.PrecoTabela, // Preenche o campo que o exportador vai pedir
            
            Ncm = "69072100",
            Cst = "060",
            
            // === CORRIGIDO PARA O SEU DTO REAL ===
            CfopDentro = "5102", 
            CfopFora = "6102",

            Cfop = "5102", // ou a regra de CFOP que você determinou para a sua planilha
            PrecoFracionado = produto.PrecoVenda, // ou a sua lógica correspondente de preço m² / fracionado
            
            PesoBruto = produto.M2Caixa,
            Espessura = produto.Espessura
        };
    }

    private static Grupo IdentificarGrupo(string tipoTabela)
    {
        if (string.IsNullOrWhiteSpace(tipoTabela)) return Grupo.NaoDefinido;
        
        return tipoTabela.ToUpper() switch
        {
            "VAREJO" => Grupo.Revestimento,
            "VINILICO" => Grupo.Vinilico,
            "LASTRA" => Grupo.Lastra,
            "VILLA_ART" => Grupo.Boutique,
            _ => Grupo.NaoDefinido
        };
    }

    private static SubGrupo IdentificarSubGrupo(string colecao, Grupo grupo)
    {
        if (grupo == Grupo.Vinilico)
        {
            if (colecao.Contains("SPC", StringComparison.OrdinalIgnoreCase)) return SubGrupo.Spc;
            if (colecao.Contains("LVT", StringComparison.OrdinalIgnoreCase)) return SubGrupo.Lvt;
            return SubGrupo.Spc;
        }

        if (grupo == Grupo.Boutique || (colecao != null && colecao.Contains("VILLA", StringComparison.OrdinalIgnoreCase))) 
            return SubGrupo.BoutiqueArt;

        return SubGrupo.Porcelanato;
    }
}