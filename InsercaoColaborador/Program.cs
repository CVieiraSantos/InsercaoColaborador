using ClosedXML.Excel;
using InsercaoColaborador.Entities;
using System.Text;
using System.Text.RegularExpressions;

string caminhoExcel = @"caminho da planilha que será salva";
string caminhoSql = @"caminho do arquivo .sql onde será salvo";

var colaboradores = new List<ColaboradorExcel>();

if (!File.Exists(caminhoExcel))
{
    Console.Error.WriteLine($"Excel file not found: {caminhoExcel}");
    return;
}

using (var fs = File.Open(caminhoExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
using (var workbook = new XLWorkbook(fs))
{
    if (workbook.Worksheets.Count == 0)
        throw new InvalidOperationException($"Workbook contains no worksheets: {caminhoExcel}");

    var planilha = workbook.Worksheet(1);
    var linhas = planilha.RangeUsed()?.RowsUsed()?.Skip(1) ?? Enumerable.Empty<IXLRangeRow>();

    foreach (var linha in linhas)
    {
        string cns = linha.Cell(1).GetString().Trim();
        string nome = linha.Cell(2).GetString().Trim();
        string cpfCnpj = Regex.Replace(linha.Cell(3).GetString(), @"\D", "");

        if (string.IsNullOrWhiteSpace(cpfCnpj) || cpfCnpj.Length != 11)
            continue;

        colaboradores.Add(new ColaboradorExcel
        {
            CNS = cns,
            Nome = nome,
            CPF = cpfCnpj
        });
    }
}

var sql = new StringBuilder();

var employee = colaboradores
    .Where(c => !string.IsNullOrWhiteSpace(c.CPF))
    .DistinctBy(c => c.CPF)
    .Select(x => new { x.Nome, x.CPF, x.CNS })
    .ToList();

if (employee.Count == 0)
{
    File.WriteAllText(caminhoSql, string.Empty, Encoding.UTF8);
    return;
}

var header = new StringBuilder();
header.AppendLine("INSERT INTO colaboradores");
header.AppendLine("(");
header.AppendLine("    IdTResponsavel,");
header.AppendLine("    IdCargo,");
header.AppendLine("    IdEntidade,");
header.AppendLine("    Nome,");
header.AppendLine("    DataNascimento,");
header.AppendLine("    Rg,");
header.AppendLine("    Cpf,");
header.AppendLine("    Endereco,");
header.AppendLine("    Numero,");
header.AppendLine("    Bairro,");
header.AppendLine("    Cidade,");
header.AppendLine("    Uf,");
header.AppendLine("    Cep,");
header.AppendLine("    TelContato1,");
header.AppendLine("    TelContato2,");
header.AppendLine("    Email1,");
header.AppendLine("    Email2,");
header.AppendLine("    DataCriacao,");
header.AppendLine("    Ativo,");
header.AppendLine("    OrgaoClasse,");
header.AppendLine("    Formacao,");
header.AppendLine("    Vinculo,");
header.AppendLine("    CargaHoraria,");
header.AppendLine("    Salario,");
header.AppendLine("    IdCliente,");
header.AppendLine("    CNS");
header.AppendLine(")");
header.AppendLine("VALUES");

var tuples = new List<string>();

foreach (var item in employee)
{
    string nomeSeguro = (item.Nome ?? ".").Trim().Replace("'", "''");
    string cnsSeguro = (item.CNS ?? ".").Trim().Replace("'", "''");
    string cpfSeguroRaw = item.CPF?.Trim() ?? "";
    string cpfDigits = Regex.Replace(cpfSeguroRaw, @"\D", "");
    string cpfSeguro = cpfDigits.Length == 11
        ? Regex.Replace(cpfDigits, @"(\d{3})(\d{3})(\d{3})(\d{2})", "$1.$2.$3-$4")
        : cpfSeguroRaw;

    var tuple = $@"(
    2, 353, 75, '{nomeSeguro}', '1900-01-01', '.', '{cpfSeguro}',
    '.', '.', '.', '.', '.', '.', NULL, NULL, NULL, NULL,
    NOW(), 1, NULL, 'Ensino Superior Completo', 'CLT', 0, 0.00, 22, '{cnsSeguro}'
)";

    tuples.Add(tuple);
}

var body = string.Join("," + Environment.NewLine, tuples);
var finalSql = header.ToString() + body + ";" + Environment.NewLine;

File.WriteAllText(caminhoSql, finalSql, Encoding.UTF8);

