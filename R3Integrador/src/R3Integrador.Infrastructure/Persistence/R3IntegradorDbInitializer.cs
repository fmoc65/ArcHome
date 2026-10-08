using System.Globalization;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace R3Integrador.Infrastructure.Persistence;

public sealed class R3IntegradorDbInitializer
{
    private static readonly ColunaProduto[] ColunasProduto =
    [
        new("CodigoInterno"), new("CodigoFabrica"), new("CodigoBarras"),
        new("DescricaoCompleta"), new("DescricaoComercial"), new("Grupo"),
        new("SubGrupo"), new("Marca"), new("Linha"), new("Modelo"),
        new("Voltagem"), new("Cor"), new("Ncm"), new("UfOrigem"),
        new("PrecoVenda", true), new("PrecoFabrica", true),
        new("DescontoPercentual", true), new("IpiPercentual", true),
        new("AliqIcmsOrigem", true), new("AliqIcmsInterna", true),
        new("Iva", true), new("FreteReais", true), new("FretePercentual", true),
        new("Unidade"), new("QtdeEmbalagemVenda", true), new("Cst"),
        new("AliquotaCofinsCst"), new("AliquotaIpiCst"), new("AliquotaPisCst"),
        new("Csosn"), new("CfopDentro"), new("CfopFora"),
        new("PesoLiquido", true), new("PesoBruto", true),
        new("QtdeEmbalagemCompra", true), new("ValorPi", true),
        new("AliquotaCofins", true), new("AliquotaPis", true),
        new("PercentualSt", true), new("UnidFabril"), new("Observacao"),
        new("DiferencaIcms", true), new("ReducaoBaseIcms", true),
        new("ReducaoBaseSt", true), new("RetencaoPis"), new("RetencaoCofins"),
        new("RetencaoCsll"), new("RetencaoIrrf"), new("RetencaoPrevSocial"),
        new("Localizacao"), new("EnquadramentoIpi"), new("AliquotaPisOrigem"),
        new("AliquotaCofinsOrigem"), new("Imagem"), new("EstoqueMinimo", true),
        new("EstoqueMaximo", true), new("AliquotaIbs"), new("AliquotaCbs"),
        new("ClassificacaoTributaria"), new("CodigoBeneficio")
    ];

    public ResultadoInicializacaoDb CriarOuAtualizar(
        string caminhoBanco,
        string caminhoPlanilha,
        string tabelaOrigem)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(caminhoBanco);
        ArgumentException.ThrowIfNullOrWhiteSpace(caminhoPlanilha);
        ArgumentException.ThrowIfNullOrWhiteSpace(tabelaOrigem);

        caminhoBanco = Path.GetFullPath(caminhoBanco);
        caminhoPlanilha = Path.GetFullPath(caminhoPlanilha);
        tabelaOrigem = tabelaOrigem.Trim().ToUpperInvariant();

        if (!File.Exists(caminhoPlanilha))
        {
            throw new FileNotFoundException("A planilha informada nao foi encontrada.", caminhoPlanilha);
        }

        var diretorioBanco = Path.GetDirectoryName(caminhoBanco);
        if (!string.IsNullOrWhiteSpace(diretorioBanco))
        {
            Directory.CreateDirectory(diretorioBanco);
        }

        using var workbook = new XLWorkbook(caminhoPlanilha);
        var worksheet = workbook.Worksheets.FirstOrDefault()
            ?? throw new InvalidDataException("A planilha nao possui abas para importar.");

        ValidarLayout(worksheet);

        using var connection = new SqliteConnection($"Data Source={caminhoBanco}");
        connection.Open();

        using var transaction = connection.BeginTransaction();
        CriarEstrutura(connection, transaction);
        RemoverCargaAnterior(connection, transaction, tabelaOrigem);
        var quantidade = ImportarProdutos(connection, transaction, worksheet, tabelaOrigem);
        ZerarEmbalagemVendaDosProdutosM2(connection, transaction, tabelaOrigem);
        transaction.Commit();

        return new ResultadoInicializacaoDb(
            caminhoBanco,
            worksheet.Name,
            tabelaOrigem,
            quantidade);
    }

    private static void ValidarLayout(IXLWorksheet worksheet)
    {
        var ultimaColuna = worksheet.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;
        if (ultimaColuna != ColunasProduto.Length)
        {
            throw new InvalidDataException(
                $"A aba {worksheet.Name} possui {ultimaColuna} colunas no cabecalho; " +
                $"eram esperadas {ColunasProduto.Length}.");
        }
    }

    private static void CriarEstrutura(SqliteConnection connection, SqliteTransaction transaction)
    {
        var colunasSql = string.Join(",\n    ", ColunasProduto.Select(coluna =>
            $"\"{coluna.Nome}\" {(coluna.Numerica ? "REAL" : "TEXT")}"));

        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = $"""
            CREATE TABLE IF NOT EXISTS "Produtos" (
                "Id" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                {colunasSql},
                "TabelaOrigem" TEXT NOT NULL
            );

            CREATE INDEX IF NOT EXISTS "IX_Produtos_TabelaOrigem"
                ON "Produtos" ("TabelaOrigem");

            CREATE TABLE IF NOT EXISTS "Usuario" (
                "Id" INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
                "Nome" TEXT NOT NULL,
                "Login" TEXT NOT NULL COLLATE NOCASE,
                "Senha" TEXT NOT NULL,
                "Ativo" INTEGER NOT NULL DEFAULT 1 CHECK ("Ativo" IN (0, 1)),
                CONSTRAINT "UQ_Usuario_Login" UNIQUE ("Login")
            );
            """;
        command.ExecuteNonQuery();
    }

    private static void RemoverCargaAnterior(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string tabelaOrigem)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = "DELETE FROM \"Produtos\" WHERE \"TabelaOrigem\" = $tabelaOrigem;";
        command.Parameters.AddWithValue("$tabelaOrigem", tabelaOrigem);
        command.ExecuteNonQuery();
    }

    private static int ImportarProdutos(
        SqliteConnection connection,
        SqliteTransaction transaction,
        IXLWorksheet worksheet,
        string tabelaOrigem)
    {
        var nomesColunas = string.Join(", ", ColunasProduto.Select(coluna => $"\"{coluna.Nome}\""));
        var parametros = string.Join(", ", Enumerable.Range(0, ColunasProduto.Length).Select(i => $"$p{i}"));

        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = $"""
            INSERT INTO "Produtos" ({nomesColunas}, "TabelaOrigem")
            VALUES ({parametros}, $tabelaOrigem);
            """;

        for (var i = 0; i < ColunasProduto.Length; i++)
        {
            command.Parameters.Add(new SqliteParameter($"$p{i}", DBNull.Value));
        }

        command.Parameters.AddWithValue("$tabelaOrigem", tabelaOrigem);

        var ultimaLinha = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        var quantidade = 0;

        for (var numeroLinha = 2; numeroLinha <= ultimaLinha; numeroLinha++)
        {
            var linha = worksheet.Row(numeroLinha);
            if (linha.Cells(1, ColunasProduto.Length).All(celula => celula.IsEmpty()))
            {
                continue;
            }

            for (var i = 0; i < ColunasProduto.Length; i++)
            {
                var celula = linha.Cell(i + 1);
                command.Parameters[i].Value = ObterValor(celula, ColunasProduto[i].Numerica);
            }

            command.ExecuteNonQuery();
            quantidade++;
        }

        return quantidade;
    }

    private static void ZerarEmbalagemVendaDosProdutosM2(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string tabelaOrigem)
    {
        // A orientação fiscal/comercial da Mônica para Nina Martinelli é
        // explícita: todas as embalagens de venda devem permanecer em 1.
        if (tabelaOrigem.Equals("NINA_MARTINELLI", StringComparison.OrdinalIgnoreCase))
        {
            return;
        }

        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            UPDATE "Produtos"
            SET "QtdeEmbalagemVenda" = 0
            WHERE "TabelaOrigem" = $tabelaOrigem
              AND UPPER(TRIM("Unidade")) = 'M2';
            """;
        command.Parameters.AddWithValue("$tabelaOrigem", tabelaOrigem);
        command.ExecuteNonQuery();
    }

    private static object ObterValor(IXLCell celula, bool numerica)
    {
        if (celula.IsEmpty())
        {
            return DBNull.Value;
        }

        if (!numerica)
        {
            return celula.GetFormattedString(CultureInfo.InvariantCulture);
        }

        if (celula.TryGetValue<double>(out var numero))
        {
            return numero;
        }

        var texto = celula.GetFormattedString(CultureInfo.InvariantCulture);
        return double.TryParse(texto, NumberStyles.Any, CultureInfo.InvariantCulture, out numero)
            ? numero
            : texto;
    }

    private sealed record ColunaProduto(string Nome, bool Numerica = false);
}

public sealed record ResultadoInicializacaoDb(
    string CaminhoBanco,
    string AbaPlanilha,
    string TabelaOrigem,
    int ProdutosImportados);
