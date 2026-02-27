using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Contrato;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.ContratoMapper;
using InsercaoColaborador.Infrastructure.Sql.Mappers.TransacaoMapper;
using InsercaoColaborador.Service;
using System.Text;

namespace InsercaoColaborador.Application.Services
{
    public class ContratoProcessamento : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Contratos";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Modelo Importacao Contratos.xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_contratos.sql";


            if (!File.Exists(caminhoExcel))
            {
                Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                return;
            }

            var contratoExcel = ExcelService.ImportarExcel(caminhoExcel, "Planilha de Contratos", linha =>
            {
                return new ContratoExcel
                {
                    Item = ValorEmInteiro.GetInt(linha.Cell(1)),
                    CnpjFornecedor = ValorEmInteiro.GetInt(linha.Cell(2)),
                    NumeroContrato = linha.Cell(3).GetString().Trim(),
                    PagamentoParcelado = ValorEmInteiro.GetInt(linha.Cell(4)),
                    QuantidadedeParcelas = ValorEmInteiro.GetInt(linha.Cell(5)),
                    TipoDeValorDoContrato = ValorEmInteiro.GetInt(linha.Cell(6)),
                    TipoDeVigencia = ValorEmInteiro.GetInt(linha.Cell(7)),
                    InicioVigencia = linha.Cell(8).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    FimVigencia = linha.Cell(9).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    DataAssinatura = linha.Cell(10).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    CriterioDeSelecao = ValorEmInteiro.GetInt(linha.Cell(11)),
                    CriterioDeSelecaoOutro = linha.Cell(12).GetString().Trim(),
                    CategoriaDeDespesa = linha.Cell(13).GetString().Trim(),
                    Valor = linha.Cell(14).GetDouble(),
                    Objeto = linha.Cell(15).GetString().Trim(),
                    NaturezaDeContratacao = linha.Cell(16).GetString().Trim(),
                    NaturezaNaoEspecificada = linha.Cell(17).GetString().Trim(),
                    ArtigoRegulamentoCompras = linha.Cell(18).GetString().Trim()
                };
            });

            var contratos = contratoExcel.Select(e =>
            {

                return new Contrato
                {
                    IdContrato = 10092,
                    NumeroContrato = e.NumeroContrato,
                    //DataTransacao = e.DataPagamento,
                    //Descricao = e.NomeCredor,
                    //Valor = e.Total,
                    //Tipo = "DEBIT",
                    //NotaFiscal = string.IsNullOrWhiteSpace(e.Documento) ? null : e.Documento,
                    //Categoria = "Outras despesas",
                    //DataNotaFiscal = e.DataPagamento,
                    //NomeBeneficiario = e.NomeCredor.Length > 100 ? e.NomeCredor.Substring(0, 100) : e.NomeCredor,
                    //OrigemRecurso = "Estadual",
                    //IdParceria = 280,
                    //Referencia = e.DataPagamento.Month,
                    //Exercicio = e.DataPagamento.Year,
                    //IdentificadorStorage = null,
                    //UrlDownloadArquivoTransacao = null,
                    //Status = 0,
                    //Avaliador = "",
                    //DataHoraAnalise = null,
                    //ValorContestado = null,
                    //Conciliado = 1,
                    //DataHoraConciliacao = DateTime.Now,
                    //ObservacoesEntidade = "",
                    //ObservacoesOrgao = "",
                    //IdUnidadeAtendimento = null,
                    //DataHoraCadastro = DateTime.Now,
                    //DataHoraUltimaAlteracao = null,
                    //IdCliente = 22,
                    //NaturezaDevolucao = null,
                    //IdEBanco = 171,
                    //MeioPagamento = 1,
                    //ValorDocumento = 0,
                    //ValorEncargos = 0,
                    //EstadoEmissor = 26,
                    //IdContrato = null,
                    //SubCategoria = 0,
                    //ItemDespesa = null,
                    //IdItemPlanoAplicacao = null,
                    //ExisteRateio = 0,
                    //PercentualRateio = null,
                    //AnaliseEscrita = "",
                    //IdRepasse = null,
                    //IdBeneficiario = null,
                    //TipoBeneficiario = null,
                };
            }).ToList();

            var sql = SqlInsertBuilder.BuildInsert(
                table: "contrato",
                columns: TransacaoSqlColumns.All,
                items: contratos,
                valuesProjection: ContratoSqlMapper.MapValues
            );

            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
        }
    }
}
