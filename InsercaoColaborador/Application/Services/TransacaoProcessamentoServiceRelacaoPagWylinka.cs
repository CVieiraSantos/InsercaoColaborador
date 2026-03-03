using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.TransacaoMapper;
using InsercaoColaborador.Service;
using System.Text;

namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceRelacaoPagWylinka : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações Relação de Pagamentos - Wylinka";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Pagamento wylinka novo.xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_wylinka.sql";


            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var transacaoExcel = ExcelService.ImportarExcel(caminhoExcel, "Relação de pagamentos", linha =>
            {
                return new TransacaoExcel
                {
                    Item = linha.Cell(1).GetString().Trim(),
                    NomeCredor = linha.Cell(2).GetString().Trim(),
                    CnpjCpf = linha.Cell(3).GetString(),
                    ServicoProduto = linha.Cell(4).GetString(),
                    TipoDespesa = linha.Cell(5).GetString(),
                    SubCategoriaDeDespesa = linha.Cell(6).GetString(),
                    Documento = linha.Cell(7).GetString().Trim(),
                    //Justificativa = linha.Cell(9).GetString(),
                    DataPagamento = linha.Cell(8).GetDateTimeOrNull()?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                    Total = linha.Cell(9).GetDecimal(),
                    
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {

                return new Transacao
                {
                    IdExtrato = 10128,
                    Numero = "",
                    DataTransacao = e.DataPagamento,
                    Descricao = string.Concat(e.TipoDespesa," - ",e.SubCategoriaDeDespesa),
                    Valor = e.Total,
                    Tipo = "DEBIT",
                    NotaFiscal = string.IsNullOrWhiteSpace(e.Documento) ? null : e.Documento,
                    Categoria = "Outras despesas",
                    DataNotaFiscal = e.DataPagamento,
                    NomeBeneficiario = e.CnpjCpf,
                    OrigemRecurso = "Estadual",
                    IdParceria = 278,
                    Referencia = e.DataPagamento.Month,
                    Exercicio = e.DataPagamento.Year,
                    IdentificadorStorage = null,
                    UrlDownloadArquivoTransacao = null,
                    Status = 1,
                    Avaliador = "João Arthur da Silva Reis",
                    DataHoraAnalise = DateTime.Now,
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
                    IdEBanco = 165,
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

            var sql = SqlInsertBuilder.BuildInsert(
                table: "transacao",
                columns: TransacaoSqlColumns.All,
                items: transacoes,
                valuesProjection: TransacaoSqlMapper.MapValues
            );
            
            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
        }
    }
}
