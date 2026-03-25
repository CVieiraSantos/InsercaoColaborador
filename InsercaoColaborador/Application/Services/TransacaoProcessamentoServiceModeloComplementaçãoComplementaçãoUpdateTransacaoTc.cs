using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using System.Text;

namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceModeloComplementaçãoUpdateTransacaoTc : IProcessamentoService
    {
        private static bool _passouPelo1218 = false;
        private static bool _gatilhoAtivado = false;
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
                var dataEmissao = ValorEmData.GetDatetime(linha.Cell(2));
                
                var rateioText = linha.Cell(8).GetString()?.Trim();
                int rateio;
                if (string.IsNullOrEmpty(rateioText))
                {
                    rateio = 0;
                }
                else if (rateioText.Equals("sim", StringComparison.OrdinalIgnoreCase) ||
                         rateioText.Equals("s", StringComparison.OrdinalIgnoreCase) ||
                         rateioText == "1")
                {
                    rateio = 1;
                }
                else if (rateioText.Equals("não", StringComparison.OrdinalIgnoreCase) ||
                         rateioText.Equals("nao", StringComparison.OrdinalIgnoreCase) ||
                         rateioText == "0")
                {
                    rateio = 0;
                }
                else
                {
                    rateio = ValorEmInteiro.GetInt(linha.Cell(8));
                }

                var percentualRateio = ValorEmDecimal.GetDecimal(linha.Cell(9));
                if (rateio == 0)
                    percentualRateio = 0m;

                var valorLiquido = ValorEmDecimal.GetDecimal(linha.Cell(5));
                var valorEncargos = ValorEmDecimal.GetDecimal(linha.Cell(6));
                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    DataEmissaoDocFiscal = dataEmissao ?? DateTime.MinValue,
                    EstadoEmissor = ValorEmInteiro.GetInt(linha.Cell(3)),
                    CnpjCpf = linha.Cell(4).GetString().Trim(),
                    ValorLiquido = valorLiquido < 0 ? tornarValorPositivo(valorLiquido) : valorLiquido,
                    ValorEncargos = valorEncargos < 0 ? tornarValorPositivo(valorEncargos) : valorEncargos,
                    SubCategoriaDeDespesa = linha.Cell(7).GetString().Trim(),
                    Rateio = rateio,
                    PercentualRateio = percentualRateio,
                    NumeroDoContrato = linha.Cell(10).GetString().Trim(),
                };
            });

            _passouPelo1218 = false;
            _gatilhoAtivado = false;

            var listaFinal = new List<Transacao>();

            foreach (var e in transacaoExcel)
            {
                var resultado = ConcatenaOuNaoConcatena(e.Item ?? string.Empty);
                if (string.IsNullOrWhiteSpace(e.Item))
                    continue;
                listaFinal.Add(new Transacao
                {
                    Numero = e.Item?.Trim() ?? string.Empty,
                    DataNotaFiscal = e.DataEmissaoDocFiscal,
                    EstadoEmissor = e.EstadoEmissor,
                    NomeBeneficiario = e.CnpjCpf.Trim() ?? string.Empty,                    
                    ValorDocumento = e.ValorLiquido == 0 ? 0m : e.ValorLiquido,
                    ValorEncargos = e.ValorEncargos == 0 ? 0m : e.ValorEncargos,
                    Categoria = e.SubCategoriaDeDespesa.Trim(),
                    ExisteRateio = (e.Rateio == 1) ? 1 : 0,
                    PercentualRateio = (e.Rateio == 1) ? e.PercentualRateio : 0m,
                    IdContrato = ObterIdContratoPorNumero(e.NumeroDoContrato),
                    ObservacoesEntidade = resultado?.Trim() ?? string.Empty,
                    IdCliente = 22,
                    IdParceria = 266
                });
            }
            var transacoes = listaFinal;
            
            var sql = SqlUpdateBuilder.BuildUpdates(
                table: "transacao",
                setProjection: TransacaoUpdateMapperComplementacao.MapUpdateSet,
                items: transacoes,
                whereProjection: TransacaoUpdateMapperComplementacao.MapWhereByAlternativeColumns
            );

            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
        }

        public static string ConcatenaOuNaoConcatena(string valorItem)
        {
            if (int.TryParse(valorItem, out int numeroItem))
            {
                if (!_gatilhoAtivado && _passouPelo1218 && numeroItem < 1218)
                {
                    _gatilhoAtivado = true;
                }

                if (numeroItem == 1218)
                {
                    _passouPelo1218 = true;
                }
            }

            return _gatilhoAtivado ? $"2.{valorItem}" : valorItem;
        }

        public int ObterIdContratoPorNumero(string numero)
        {
            if (string.IsNullOrEmpty(numero))
                return 0;

            var contratos = new List<KeyValuePair<string, int>>
            {
                new KeyValuePair<string, int>("26125", 3),
                new KeyValuePair<string, int>("26131", 4),
                new KeyValuePair<string, int>("26148", 5),
                new KeyValuePair<string, int>("26129", 6),
                new KeyValuePair<string, int>("26132", 7),
                new KeyValuePair<string, int>("26139", 8),
                new KeyValuePair<string, int>("26142", 9),
                new KeyValuePair<string, int>("26145", 10),
                new KeyValuePair<string, int>("26154", 11),
                new KeyValuePair<string, int>("26156", 12),
                new KeyValuePair<string, int>("26178", 13),
                new KeyValuePair<string, int>("26216", 14),
                new KeyValuePair<string, int>("26183", 15),
                new KeyValuePair<string, int>("26185", 16),
                new KeyValuePair<string, int>("26174", 17),
                new KeyValuePair<string, int>("26175", 18),
                new KeyValuePair<string, int>("27211", 19),
                new KeyValuePair<string, int>("26180", 20),
                new KeyValuePair<string, int>("27212", 21),
                new KeyValuePair<string, int>("26215", 22),
                new KeyValuePair<string, int>("34453", 23),
                new KeyValuePair<string, int>("29829", 24),
                new KeyValuePair<string, int>("26460", 25),
                new KeyValuePair<string, int>("26459", 26),
                new KeyValuePair<string, int>("26465", 27),
                new KeyValuePair<string, int>("26462", 28),
                new KeyValuePair<string, int>("26461", 29),
                new KeyValuePair<string, int>("26463", 30),
                new KeyValuePair<string, int>("26464", 31),
                new KeyValuePair<string, int>("26466", 32),
                new KeyValuePair<string, int>("26468", 33),
                new KeyValuePair<string, int>("26467", 34),
                new KeyValuePair<string, int>("26470", 35),
                new KeyValuePair<string, int>("26958", 36),
                new KeyValuePair<string, int>("26968", 37),
                new KeyValuePair<string, int>("26182", 38),
                new KeyValuePair<string, int>("26176", 39),
                new KeyValuePair<string, int>("28936", 40),
                new KeyValuePair<string, int>("26439", 41),
                new KeyValuePair<string, int>("30633", 42),
                new KeyValuePair<string, int>("26440", 43),
                new KeyValuePair<string, int>("26471", 44),
                new KeyValuePair<string, int>("26472", 45),
                new KeyValuePair<string, int>("26473", 46),
            };

            var resultado = contratos.FirstOrDefault(x => x.Key == numero);
            return resultado.Value;
        }

        public decimal tornarValorPositivo(decimal valor)
        {
            return valor * -1;
        }
    }
}
