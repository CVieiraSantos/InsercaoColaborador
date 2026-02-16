using ClosedXML.Excel;
using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Service;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceRelacaoPagConvenio : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações Relação de Pagamentos - Assinado (1)";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Relação de Pagamentos - Assinado (1).xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_novo_Consolidação de Pagamentos e Es.sql";


            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var transacaoExcel = ExcelService.ImportarExcel(caminhoExcel, "Consolidação de Pagamentos e Es", linha =>
            {
                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    NomeCredor = linha.Cell(2).GetString().Trim(),
                    Documento = linha.Cell(3).GetString().Trim(),
                    NumeroCheque = linha.Cell(4).GetString().Trim(),
                    //DataPagamento = ParseExcelDate(linha.Cell(5)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                    DataPagamento = linha.Cell(5).GetDateTimeOrNull()?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                    Total = linha.Cell(6).GetDecimal(),
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {
                //var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);

                return new Transacao
                {
                    IdExtrato = 10092,
                    Numero = e.NumeroCheque,
                    DataTransacao = e.DataPagamento,
                    Descricao = e.NomeCredor,
                    Valor = e.Total,
                    Tipo = "DEBIT",
                    NotaFiscal = string.IsNullOrWhiteSpace(e.Documento) ? null : e.Documento,
                    Categoria = "Outras despesas",
                    DataNotaFiscal = e.DataPagamento,
                    //NomeBeneficiario = nomeBeneficiario.Length > 100 ? nomeBeneficiario.Substring(0, 100) : nomeBeneficiario,
                    NomeBeneficiario = e.NomeCredor.Length > 100 ? e.NomeCredor.Substring(0,100) : e.NomeCredor,
                    OrigemRecurso = "Estadual",
                    IdParceria = 280,
                    Referencia = e.DataPagamento.Month,
                    Exercicio = e.DataPagamento.Year,
                    IdentificadorStorage = null,
                    UrlDownloadArquivoTransacao = null,
                    Status = 0,
                    Avaliador = "",
                    DataHoraAnalise = null,
                    ValorContestado = null,
                    Conciliado = 1,
                    DataHoraConciliacao = DateTime.Now,
                    ObservacoesEntidade = "",
                    ObservacoesOrgao = "",
                    IdUnidadeAtendimento = null,
                    DataHoraCadastro = DateTime.Now,
                    DataHoraUltimaAlteracao = null,
                    IdCliente = 22,
                    NaturezaDevolucao = null,
                    IdEBanco = 171,
                    MeioPagamento = 1,
                    ValorDocumento = 0,
                    ValorEncargos = 0,
                    EstadoEmissor = 26,
                    IdContrato = null,
                    SubCategoria = 0,
                    ItemDespesa = null,
                    IdItemPlanoAplicacao = null,
                    ExisteRateio = 0,
                    PercentualRateio = null,
                    AnaliseEscrita = "",
                    IdRepasse = null,
                    IdBeneficiario = null,
                    TipoBeneficiario = null,
                };
            }).ToList();

            var header = new StringBuilder();

            header.AppendLine("INSERT INTO transacao");
            header.AppendLine("(");
            header.AppendLine("    IdExtrato, Numero, DataTransacao, Descricao, Valor, Tipo, NotaFiscal, Categoria, DataNotaFiscal,");
            header.AppendLine("    NomeBeneficiario, OrigemRecurso, IdParceria, Referencia, Exercicio, Status, Avaliador,");
            header.AppendLine("    ValorContestado, Conciliado, DataHoraConciliacao, ObservacoesEntidade, ObservacoesOrgao,");
            header.AppendLine("    DataHoraCadastro, IdCliente, IdEBanco, MeioPagamento, ValorDocumento, ValorEncargos,");
            header.AppendLine("    EstadoEmissor, SubCategoria, ExisteRateio, AnaliseEscrita");
            header.AppendLine(") VALUES");

            var tuples = transacoes.Select(item =>
            $@"(
                {SqlInt(item.IdExtrato)},
                {SqlString(item.Numero)},
                {SqlDateTime(item.DataTransacao)},
                {SqlString(item.Descricao)},
                {SqlDecimal(item.Valor)},
                {SqlString(item.Tipo)},
                {SqlString(item.NotaFiscal)},
                {SqlString(item.Categoria)},
                {SqlDateTime(item.DataNotaFiscal)},
                {SqlString(item.NomeBeneficiario)},
                {SqlString(item.OrigemRecurso)},
                {SqlInt(item.IdParceria)},
                {SqlInt(item.Referencia)},
                {SqlInt(item.Exercicio)},
                {SqlInt(item.Status)},
                {SqlString(item.Avaliador)},
                {SqlDecimalNullable(item.ValorContestado)},
                {SqlInt(item.Conciliado)},
                {SqlDateTimeNullable(item.DataHoraConciliacao)},
                {SqlString(item.ObservacoesEntidade)},
                {SqlString(item.ObservacoesOrgao)},
                {SqlDateTimeNullable(item.DataHoraCadastro)},
                {SqlInt(item.IdCliente)},
                {SqlIntNullable(item.IdEBanco)},
                {SqlInt(item.MeioPagamento)},
                {SqlDecimal(item.ValorDocumento)},
                {SqlDecimal(item.ValorEncargos)},
                {SqlInt(item.EstadoEmissor)},
                {SqlInt(item.SubCategoria)},
                {SqlInt(item.ExisteRateio)},
                {SqlString(item.AnaliseEscrita)}
            )");

            File.WriteAllText(caminhoSql, header.ToString() + string.Join("," + Environment.NewLine, tuples) + ";", Encoding.UTF8);

            static string SqlString(string? s) => s == null ? "NULL" : $"'{s.Replace("'", "''")}'";
            static string SqlDecimal(decimal d) => d.ToString("F2", CultureInfo.InvariantCulture);
            static string SqlDecimalNullable(decimal? d) => d.HasValue ? SqlDecimal(d.Value) : "NULL";
            static string SqlInt(int i) => i.ToString(CultureInfo.InvariantCulture);
            static string SqlIntNullable(int? i) => i.HasValue ? SqlInt(i.Value) : "NULL";
            static string SqlDateTime(DateTime dt) => $"'{dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}'";
            static string SqlDateTimeNullable(DateTime? dt) => dt.HasValue ? SqlDateTime(dt.Value) : "NULL";

            //DateTime? ParseExcelDate(IXLCell cell)
            //{
            //    if (cell.Value.IsDateTime) 
            //        return cell.GetDateTime();

            //    var s = cell.GetString()?.Trim();
                
            //    if (string.IsNullOrEmpty(s)) 
            //        return null;

            //    if (DateTime.TryParse(s, new CultureInfo("pt-BR"), DateTimeStyles.None, out DateTime dt))
            //        return dt;

            //    return null;
            //}


        }
    }
}
