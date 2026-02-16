using ClosedXML.Excel;
using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Service;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServicePendencias1 : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações Pendencias 1";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\CPFs para montagem de insert(2).xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_novo_pendencias1.sql";


            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var transacaoExcel = ExcelService.ImportarExcel(caminhoExcel, "pendencias 1", linha =>
            {
                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    NomeCredor = linha.Cell(2).GetString().Trim(),
                    CnpjCpf = CpfCnpjGenerator.FormatarCpfOuCnpj(linha.Cell(3).GetString()),
                    NotaFiscalOuEquivalente = linha.Cell(4).GetString().Trim(),
                    ServicoProduto = linha.Cell(5).GetString().Trim(),
                    NumeroCheque = linha.Cell(6).GetString().Trim(),
                    DataPagamento = ParseExcelDate(linha.Cell(7)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                    Total = linha.Cell(8).GetDecimal(),
                    ValorGlosado = linha.Cell(9).IsEmpty() ? null : (decimal?)linha.Cell(9).GetDecimal(),
                    ValorConciliado = linha.Cell(11).GetString().Trim().Equals("SIM", StringComparison.OrdinalIgnoreCase),
                    StatusAnalise = linha.Cell(12).GetString().Trim(),
                    StatusCor = GetCellBackgroundColorKey(linha.Cell(12)),
                    Observacoes = linha.Cell(13).GetString().Trim()
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {
                var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);

                return new Transacao
                {
                    IdExtrato = 10092,
                    Numero = e.NumeroCheque,
                    DataTransacao = e.DataPagamento,
                    Descricao = e.ServicoProduto.Length > 200 ? e.ServicoProduto.Substring(0, 200) : e.ServicoProduto,
                    Valor = e.Total,
                    Tipo = "DEBIT",
                    NotaFiscal = string.IsNullOrWhiteSpace(e.NotaFiscalOuEquivalente) ? null : e.NotaFiscalOuEquivalente,
                    Categoria = "Outras despesas",
                    DataNotaFiscal = e.DataPagamento,
                    NomeBeneficiario = nomeBeneficiario.Length > 100 ? nomeBeneficiario.Substring(0, 100) : nomeBeneficiario,
                    OrigemRecurso = "Estadual",
                    IdParceria = 266,
                    Referencia = e.DataPagamento.Month,
                    Exercicio = e.DataPagamento.Year,
                    Status = ResolverStatus(e.StatusAnalise, e.StatusCor),
                    Avaliador = "Maria Cristina Figueiredo Shigaki",
                    ValorContestado = e.ValorGlosado,
                    Conciliado = e.ValorConciliado ? 1 : 0,
                    DataHoraConciliacao = DateTime.Now,
                    ObservacoesEntidade = e.Item,
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
                    AnaliseEscrita = e.StatusAnalise
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

            static int ResolverStatus(string statusAnalise, string? statusCor)
            {

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

                if (string.IsNullOrWhiteSpace(statusAnalise))
                    return 0;

                var s = statusAnalise.Trim().ToLowerInvariant();
                return s switch
                {
                    "regular" => 1,
                    "irregular" => 2,
                    "regular com ressalvas" => 3,

                    _ => 0
                };
            }

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
                //if (cell.Value.IsDateTime) return cell.GetDateTime();

                var s = cell.GetString()?.Trim();
                if (string.IsNullOrEmpty(s)) return null;

                var dataString = s.Split("/"); // 1/13/2025
                                               // 10/1/2025

                int.TryParse(dataString[0], out int dia);

                int.TryParse(dataString[1], out int mes);

                if (dia > 0 && dia < 13 && mes <= 12)
                {
                    s = string.Concat(dataString[1], "/", dataString[0], "/", dataString[2]);
                    if (DateTime.TryParse(s, new CultureInfo("pt-BR"), DateTimeStyles.None, out DateTime dta))
                        return dta;
                }

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

        }
    }
}
