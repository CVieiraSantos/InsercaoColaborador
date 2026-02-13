using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Colaborador;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Service;
using System.Globalization;
using System.Text;

namespace InsercaoColaborador.Application.Services
{
    public class ColaboradorProcessamentoService : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Colaboradores";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\CPFs para montagem de insert(1).xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_colaboradores_novo.sql";


            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var colaboradoresExcel = ExcelService.ImportarExcel(caminhoExcel, "Página1", linha =>
            {
                return new ColaboradorExcel
                {
                    ItemExcel = linha.Cell(1).GetString().Trim(),
                    CredorExcel = linha.Cell(2).GetString().Trim(),
                    CPFExcel = linha.Cell(3).GetString().Trim()
                };
            });

            var colaboradores =
            colaboradoresExcel
            .Select
            (c =>
                new Colaborador
                {
                    Nome = c.CredorExcel,
                    CPF = c.CPFExcel,
                    CNS = c.ItemExcel
                }
            );

            var header = new StringBuilder();

            var employee = colaboradores
                .Where(c => !string.IsNullOrWhiteSpace(CpfCnpjGenerator.FormatarCpf(c)))
                .DistinctBy(c => c.CPF)
                .Select(x => new { x.Nome, x.CPF, x.CNS })
                .ToList();


            if (employee.Count == 0)
            {
                File.WriteAllText(caminhoSql, string.Empty, Encoding.UTF8);
                return;
            }

            header.Clear();
            header.AppendLine("INSERT INTO colaboradores");
            header.AppendLine("(");
            header.AppendLine("    IdTResponsavel, IdCargo, IdEntidade, Nome, DataNascimento, Rg, Cpf, Endereco,");
            header.AppendLine("    Numero, Bairro, Cidade, Uf, Cep, TelContato1, TelContato2, Email1, Email2,");
            header.AppendLine("    DataCriacao, Ativo, OrgaoClasse, Formacao, Vinculo, CargaHoraria, Salario,");
            header.AppendLine("    IdCliente, CNS");
            header.AppendLine(") VALUES");

            var tuples = employee.Select(item =>
            $@"(
                {SqlInt(2)},
                {SqlInt(353)},
                {SqlInt(75)},
                {SqlString(item.Nome)},
                {SqlDateTime(DateTime.Now)},
                '.',
                {SqlString(CpfCnpjGenerator.FormatarCpf(new Colaborador { CPF = item.CPF }))},
                '.',
                '.',
                '.',
                '.',
                '.',
                '.',
                NULL,
                NULL,
                NULL,
                NULL,
                {SqlDateTime(DateTime.Now)},
                {SqlInt(1)},
                NULL,
                {SqlString("Ensino Superior Completo")},
                {SqlString("CLT")},
                {SqlInt(0)},
                {SqlDecimal(0m)},
                {SqlInt(22)},
                {SqlString(item.CNS)}
            )");

            File.WriteAllText(caminhoSql, header.ToString() + string.Join("," + Environment.NewLine, tuples) + ";", Encoding.UTF8);

            static string SqlString(string? s) => s == null ? "NULL" : $"'{s.Replace("'", "''")}'";
            static string SqlDecimal(decimal d) => d.ToString("F2", CultureInfo.InvariantCulture);
            static string SqlInt(int i) => i.ToString(CultureInfo.InvariantCulture);
            static string SqlDateTime(DateTime dt) => $"'{dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}'";
        }
    }
}
