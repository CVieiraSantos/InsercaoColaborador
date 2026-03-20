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
        private bool _passouPelo1218 = false;
        private bool _gatilhoAtivado = false;
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
                var dataEmissao = ValorEmData.GetDatetime(linha.Cell(4));
                if (!dataEmissao.HasValue)
                {
                    Console.Error.WriteLine($"Linha {linha.RowNumber()}: Data inválida em coluna 4 ('{linha.Cell(4).GetString()}'). Linha pulada.");
                    return null;
                }

                var rateioText = linha.Cell(12).GetString()?.Trim();
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
                    rateio = ValorEmInteiro.GetInt(linha.Cell(12));
                }

                var percentualRateio = ValorEmDecimal.GetDecimal(linha.Cell(13));
                if (rateio == 0)
                    percentualRateio = 0m;


                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    NomeCredor = linha.Cell(2).GetString().Trim(),
                    NotaFiscalOuEquivalente = linha.Cell(3).GetString().Trim(),
                    DataEmissaoDocFiscal = dataEmissao.Value,
                    EstadoEmissor = ValorEmInteiro.GetInt(linha.Cell(5)),
                    EstadoEmissorDescricao = linha.Cell(6).GetString().Trim(),
                    CnpjCpf = CpfCnpjGenerator.FormatarCpfOuCnpj(linha.Cell(7).GetString().Trim()),
                    ValorBruto = ValorEmDecimal.GetDecimal(linha.Cell(8)),
                    ValorLiquido = ValorEmDecimal.GetDecimal(linha.Cell(9)),
                    ValorEncargos = ValorEmDecimal.GetDecimal(linha.Cell(10)),
                    SubCategoriaDeDespesa = linha.Cell(11).GetString().Trim(),
                    Rateio = rateio,
                    PercentualRateio = percentualRateio,
                    NumeroDoContrato = linha.Cell(14).GetString().Trim(),
                };
            });

            _passouPelo1218 = false;
            _gatilhoAtivado = false;

            var listaFinal = new List<Transacao>();

            foreach (var e in transacaoExcel)
            {
                var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);
                var resultado = ConcatenaOuNaoConcatena(e.Item ?? string.Empty);
                
                listaFinal.Add(new Transacao
                {
                    Numero = e.Item?.Trim() ?? string.Empty,
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
                    ObservacoesEntidade = resultado?.Trim() ?? string.Empty,
                    IdCliente = 22,
                    IdParceria = 0

                });
            }
            var transacoes = listaFinal;
            
            var sql = SqlUpdateBuilder.BuildUpdates(
                table: "transacao",
                setProjection: TransacaoUpdateMapperComplementacao.MapUpdateSet,
                items: transacoes,
                whereProjection: TransacaoUpdateMapperComplementacao.MapWhereByAlternativeColumns // ou MapWhereByIdExtrato
            );

            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
        }

        public string ConcatenaOuNaoConcatena(string valorItem)
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
    }
}
