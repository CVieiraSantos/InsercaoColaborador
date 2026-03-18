using DocumentFormat.OpenXml.Bibliography;
using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Service;
using System.Text;


namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceModeloComplementaçãoUpdateTransacaoTc : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações Complementação - update transacao tc2024";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Complementacao update transacao tc2024.xlsx";
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
                    NomeCredor = linha.Cell(2).GetString().Trim(),
                    NotaFiscalOuEquivalente = linha.Cell(3).GetString().Trim(),
                    DataEmissaoDocFiscal = ValorEmData.GetDatetime(linha.Cell(4)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    EstadoEmissor = ValorEmInteiro.GetInt(linha.Cell(5)),
                    EstadoEmissorDescricao = linha.Cell(6).GetString().Trim(),
                    CnpjCpf = CpfCnpjGenerator.FormatarCpfOuCnpj(linha.Cell(7).GetString().Trim()),
                    ValorBruto = ValorEmDecimal.GetDecimal(linha.Cell(8)),
                    ValorLiquido = ValorEmDecimal.GetDecimal(linha.Cell(9)),
                    ValorEncargos = ValorEmDecimal.GetDecimal(linha.Cell(10)),
                    SubCategoriaDeDespesa = linha.Cell(11).GetString().Trim(),
                    Rateio = ValorEmInteiro.GetInt(linha.Cell(12)),
                    PercentualRateio = ValorEmDecimal.GetDecimal(linha.Cell(13)),
                    NumeroDoContrato = linha.Cell(14).GetString().Trim(),
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {
                var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);
                int.TryParse(e.Item, out int numeroItem);
                var itemFormatado = (numeroItem > 1218) ? $"2.{e.Item}" : e.Item;

                return new Transacao
                {
                    Numero = itemFormatado?.Trim() ?? string.Empty,
                    NomeBeneficiario = nomeBeneficiario.Length > 100 ? nomeBeneficiario.Substring(0, 100) : nomeBeneficiario,
                    NotaFiscal = e.NotaFiscalOuEquivalente,
                    DataNotaFiscal = e.DataEmissaoDocFiscal,
                    EstadoEmissor = e.EstadoEmissor,
                    EstadoEmissorDescricao = e.EstadoEmissorDescricao.Trim(),
                    Valor = e.ValorBruto,
                    ValorDocumento = e.ValorLiquido == 0 ? 0m : e.ValorLiquido,
                    ValorEncargos = e.ValorEncargos == 0 ? 0m : e.ValorEncargos,
                    Categoria = e.SubCategoriaDeDespesa.Trim(),
                    ExisteRateio = (e.Rateio == 1) ? 1 : 0,
                    PercentualRateio = (e.Rateio == 1) ? e.PercentualRateio : 0m,
                    NumeroDoContrato = e.NumeroDoContrato.Trim(),
                    ObservacoesEntidade = itemFormatado?.Trim() ?? string.Empty,
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
