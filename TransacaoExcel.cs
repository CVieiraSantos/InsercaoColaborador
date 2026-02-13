using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Drawing;
using ClosedXML.Excel;
using InsercaoColaborador.Entities;
using InsercaoColaborador.Extension;

// --- Configurações de Caminho ---
string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\CPFs para montagem de insert(1).xlsx";
string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes.sql";

if (!File.Exists(caminhoExcel))
{
    Console.Error.WriteLine($"Arquivo Excel não encontrado!");
    return;
}

// --- 1. Importação Corrigida conforme a Imagem da Planilha ---
var transacaoExcel = ExcelService.ImportarExcel(caminhoExcel, "pendencias 1", linha =>
{
    return new TransacaoExcel
    {
        Item = linha.Cell(1).GetString().Trim(),
        NomeCredor = linha.Cell(2).GetString().Trim(),
        CnpjCpf = Regex.Replace(linha.Cell(3).GetString(), @"\D", ""), // Coluna C
        NotaFiscalOuEquivalente = linha.Cell(4).GetString().Trim(),    // Coluna D
        ServicoProduto = linha.Cell(5).GetString().Trim(),             // Coluna E
        NumeroCheque = linha.Cell(6).GetString().Trim(),               // Coluna F
        DataPagamento = ParseExcelDate(linha.Cell(7)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),                   // Coluna G
        Total = linha.Cell(8).GetDecimal(),                            // Coluna H
        ValorGlosado = linha.Cell(9).IsEmpty() ? null : (decimal?)linha.Cell(9).GetDecimal(), // Coluna I
        ValorConciliado = linha.Cell(11).GetString().Trim().Equals("SIM", StringComparison.OrdinalIgnoreCase), // Coluna K
        StatusAnalise = linha.Cell(12).GetString().Trim(),             // Coluna L (texto)
        // Captura uma chave simples da cor de fundo para priorizar por cor
        StatusCor = GetCellBackgroundColorKey(linha.Cell(12)),
        Observacoes = linha.Cell(13).GetString().Trim()                // Coluna M
    };
});

// --- 2. Conversão e Regras do Tech Lead ---
var transacoes = transacaoExcel.Select(e =>
{
    var nomeBeneficiario = (e.NomeCredor?.Contains('-') ?? false) ? e.NomeCredor : e.CnpjCpf;

    return new Transacao
    {
        Numero = e.NumeroCheque,
        DataTransacao = e.DataPagamento,
        Descricao = e.ServicoProduto,
        Valor = e.Total,
        Tipo = "DEBIT",
        NotaFiscal = string.IsNullOrWhiteSpace(e.NotaFiscalOuEquivalente) ? null : e.NotaFiscalOuEquivalente,
        Categoria = "Outras despesas",
        DataNotaFiscal = e.DataPagamento,
        NomeBeneficiario = nomeBeneficiario,
        OrigemRecurso = "Estadual",
        IdParceria = 266,
        Referencia = e.DataPagamento.Month,
        Exercicio = e.DataPagamento.Year,
        // Agora passa texto + cor (cor tem prioridade)
        Status = ResolverStatus(e.StatusAnalise, e.StatusCor),
        Avaliador = "Maria Cristina Figueiredo Shigaki",
        ValorContestado = e.ValorGlosado,
        Conciliado = e.ValorConciliado ? 1 : 0,
        DataHoraConciliacao = DateTime.Now,
        ObservacoesEntidade = e.Item, // Conforme pedido: Coluna Item (A)
        ObservacoesOrgao = e.Observacoes,
        DataHoraCadastro = DateTime.Now,
        IdCliente = 22,
        IdEBanco = 161,
        MeioPagamento = 1,
        ValorDocumento = 0,
        ValorEncargos = 0,
        EstadoEmissor = 26,
        SubCategoria = 0,
        ExisteRateio = 0,
        AnaliseEscrita = ""
    };
}).ToList();

// ... (resto do arquivo permanece igual: geração SQL e helpers) ...

static int ResolverStatus(string statusAnalise, string? statusCor)
{
    // 1) Priorizar cor se presente
    if (!string.IsNullOrWhiteSpace(statusCor))
    {
        switch (statusCor.Trim().ToLowerInvariant())
        {
            case "green":
                return 1; // verde -> 1
            case "blue":
                return 3; // azul -> 3
            case "red":
                return 2; // vermelho -> 2
        }
    }

    // 2) Fallback para texto da coluna (case-insensitive)
    if (string.IsNullOrWhiteSpace(statusAnalise)) return 2;
    var s = statusAnalise.Trim().ToLowerInvariant();
    return s switch
    {
        "regular" => 1,
        "ressalvas" => 2,
        "irregular" => 3,
        _ => 2
    };
}

/// <summary>
/// Detecta uma chave simples da cor de fundo: "green","blue","red" ou null.
/// Usa heurística sobre o nome conhecido da cor ou HTML/ARGB quando necessário.
/// </summary>
static string? GetCellBackgroundColorKey(IXLCell cell)
{
    try
    {
        var xlColor = cell.Style.Fill.BackgroundColor;
        // Se não há cor, retorna null
        if (xlColor.Color.IsEmpty) return null;

        var c = xlColor.Color;
        var name = c.IsKnownColor ? c.Name.ToLowerInvariant() : ColorTranslator.ToHtml(c).ToLowerInvariant();

        if (name.Contains("green")) return "green";
        if (name.Contains("blue")) return "blue";
        if (name.Contains("red")) return "red";

        return null;
    }
    catch
    {
        // Se algo falhar ao acessar a cor (compatibilidade / plataforma), ignore e retorne null
        return null;
    }
}

DateTime? ParseExcelDate(IXLCell cell)
{
    if (cell.Value.IsDateTime) return cell.GetDateTime();

    var s = cell.GetString()?.Trim();
    if (string.IsNullOrEmpty(s)) return null;

    // Try common formats and cultures
    var formats = new[] { "M/d/yyyy", "M/d/yy", "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };
    if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
        return dt;
    if (DateTime.TryParse(s, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
        return dt;
    if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
        return dt;

    // Last resort: if numeric-looking, try OADate
    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
        return DateTime.FromOADate(d);

    return null;
}