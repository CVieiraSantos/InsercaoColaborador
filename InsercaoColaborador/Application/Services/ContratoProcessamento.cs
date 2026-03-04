using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Contrato;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.ContratoMapper;
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
                    CnpjFornecedor = linha.Cell(2).GetString().Trim(),
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
                    IdFornecedor = null,
                    RazaoSocialFornecedor = e.CnpjFornecedor.Trim(),
                    NumeroContrato = e.NumeroContrato.Trim(),
                    Parcelado = e.PagamentoParcelado,
                    QuantidadedeParcelas = e.QuantidadedeParcelas,
                    TipoValorContrato = e.TipoDeValorDoContrato,
                    TipoVigencia = e.TipoDeVigencia,
                    Inicio = e.InicioVigencia,
                    Fim = e.FimVigencia,
                    DataAssinatura = e.DataAssinatura,
                    CriterioSelecao = e.CriterioDeSelecao,
                    CriterioSelecaoOutro = e.CriterioDeSelecao == 4 ? e.CriterioDeSelecaoOutro : string.Empty,
                    CategoriaDespesa = e.CategoriaDeDespesa.Trim(),
                    Valor = e.Valor,
                    ObjetoContrato = e.Objeto,
                    NaturezaContratacao = e.NaturezaDeContratacao,
                    NaturezaContratacaoOutro = e.NaturezaDeContratacao == "23"? e.NaturezaNaoEspecificada : string.Empty,
                    ArtigoRegulamentoCompras = e.ArtigoRegulamentoCompras,
                    Ativo = 1,
                    IdCliente = 22
                };
            }).ToList();

            var sql = SqlInsertBuilder.BuildInsert(
                table: "contrato",
                columns: ContratoSqlColumns.All,
                items: contratos,
                valuesProjection: ContratoSqlMapper.MapValues
            );

            File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
        }
    }
}
