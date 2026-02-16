using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.TransacaoMapper;
using InsercaoColaborador.Service;
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
                    DataPagamento = linha.Cell(5).GetDateTimeOrNull()?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                    Total = linha.Cell(6).GetDecimal(),
                };
            });

            var transacoes = transacaoExcel.Select(e =>
            {

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
