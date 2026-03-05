using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Service;
using System.Text;


namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceModeloComplementaçãoDespesasPagas : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações Modelo de Complementação Despesas Pagas";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Modelo de Complementação Despesas Pagas.xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_novo_Complementação Pagamentos.sql";

            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var transacaoExcel = ExcelService.ImportarExcel(caminhoExcel, "Complementação Pagamentos", linha =>
            {
                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    NotaFiscalOuEquivalente = linha.Cell(2).GetString().Trim(),
                    DataEmissaoDocFiscal = ValorEmData.GetDatetime(linha.Cell(3)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    EstadoEmissor = ValorEmInteiro.GetInt(linha.Cell(4)),
                    CnpjCpf = linha.Cell(5).GetString().Trim(),
                    ValorBruto = ValorEmDecimal.GetDecimal(linha.Cell(6)),
                    ValorEncargos = ValorEmDecimal.GetDecimal(linha.Cell(7)),
                    SubCategoriaDeDespesa = linha.Cell(8).GetString().Trim(),
                    Rateio = ValorEmInteiro.GetInt(linha.Cell(9)),
                    PercentualRateio = ValorEmDecimal.GetDecimal(linha.Cell(10)),
                    NumeroDoContrato = linha.Cell(11).GetString().Trim(),
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {
                var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);

                return new Transacao
                {
                    Numero = e.Item?.Trim() ?? string.Empty,
                    NotaFiscal = e.NotaFiscalOuEquivalente,
                    DataNotaFiscal = e.DataEmissaoDocFiscal,
                    EstadoEmissor = e.EstadoEmissor,
                    NomeBeneficiario = e.CnpjCpf.Trim(),
                    Valor = e.ValorBruto,
                    ValorEncargos = e.ValorEncargos,
                    Categoria = e.SubCategoriaDeDespesa.Trim(),
                    ExisteRateio = e.Rateio ?? 0,
                    PercentualRateio = e.Rateio > 0 ? e.PercentualRateio : null,
                    NumeroDoContrato = e.NumeroDoContrato.Trim(),
                    ObservacoesEntidade = "2",
                    IdCliente = 22,
                    IdParceria = 0
                };
            }).ToList();

            var sql = SqlUpdateBuilder.BuildUpdates(
                table: "transacao",
                setProjection: TransacaoUpdateMapperComplementacao.MapUpdateSet,
                items: transacoes,
                whereProjection: TransacaoUpdateMapperComplementacao.MapWhereByAlternativeColumns // ou MapWhereByIdExtrato
            );

            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);

            //DateTime? ParseExcelDate(IXLCell cell)
            //{

            //    var s = cell.GetString()?.Trim();
            //    if (string.IsNullOrEmpty(s)) return null;

            //    if (DateTime.TryParse(s, new CultureInfo("pt-BR"), DateTimeStyles.None, out DateTime dta))
            //        return dta;

            //    var formats = new[] { "M/d/yyyy", "M/d/yy", "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };
            //    if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            //        return dt;
            //    if (DateTime.TryParse(s, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
            //        return dt;
            //    if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
            //        return dt;

            //    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
            //        return DateTime.FromOADate(d);

            //    return null;
            //}
        }
    }
}
