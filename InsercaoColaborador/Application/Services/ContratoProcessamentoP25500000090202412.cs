using DocumentFormat.OpenXml.Office2010.Excel;
using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Contrato;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.ContratoMapper;
using InsercaoColaborador.Service;
using System.Text;
using System.Text.RegularExpressions;

namespace InsercaoColaborador.Application.Services
{
    public class ContratoProcessamentoP25500000090202412 : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Contratos";

        public void Executar()
        {
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Modelo Importacao Contratos P25500000090202412.xlsx";
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
                    //Item = ValorEmInteiro.GetInt(linha.Cell(1)),
                    //NomeBeneficiario = linha.Cell(1).GetString(),
                    CnpjFornecedor = linha.Cell(1).GetString().Trim(),
                    NumeroContrato = linha.Cell(2).GetString().Trim(),
                    PagamentoParcelado = ValorEmInteiro.GetInt(linha.Cell(3)),
                    QuantidadedeParcelas = ValorEmInteiro.GetInt(linha.Cell(4)),
                    TipoDeValorDoContrato = ValorEmInteiro.GetInt(linha.Cell(5)),
                    TipoDeVigencia = ValorEmInteiro.GetInt(linha.Cell(6)),
                    InicioVigencia = linha.Cell(7).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    FimVigencia = linha.Cell(8).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    DataAssinatura = linha.Cell(9).GetDateTimeOrNull() ?? throw new FormatException($"Invalid data in row {linha.RowNumber()}"),
                    CriterioDeSelecao = ValorEmInteiro.GetInt(linha.Cell(10)),
                    CriterioDeSelecaoOutro = linha.Cell(11).GetString().Trim(),
                    CategoriaDeDespesa = linha.Cell(12).GetString().Trim(),
                    Valor = linha.Cell(13).GetDouble(),
                    Objeto = linha.Cell(14).GetString(),
                    NaturezaDeContratacao = linha.Cell(15).GetString().Trim(),
                    NaturezaNaoEspecificada = linha.Cell(16).GetString().Trim(),
                    ArtigoRegulamentoCompras = linha.Cell(17).GetString().Trim()
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
                    NaturezaContratacaoOutro = e.NaturezaDeContratacao == "23" ? e.NaturezaNaoEspecificada : string.Empty,
                    ArtigoRegulamentoCompras = e.ArtigoRegulamentoCompras,
                    TipoFornecedor = 1
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

        public int ObterIdPorCnpjPorFornecedor(string cnpj)
        {
            IDictionary<string,int> contratoFornecedores = new Dictionary<string, int>
            {
                ["48.549.331/0001-01"] = 287,
                ["02.287.620/0001-89"] = 288,
                ["44.166.627/0001-92"] = 289,
                ["03.858.331/0001-55"] = 290,
                ["23.344.957/0001-50"] = 291,
                ["04.436.106/0001-93"] = 292,
                ["61.695.227/0001-93"] = 293,
                ["08.306.796/0001-17"] = 294,
                ["23.236.675/0001-30"] = 295,
                ["22.314.093/0001-61"] = 296,
                ["32.643.177/0001-00"] = 297,
                ["23.035.684/0001-62"] = 298,
                ["45.728.775/0001-16"] = 299,
                ["09.294.688/0001-34"] = 300,
                ["29.016.530/0001-00"] = 301,
                ["03.252.307/0001-78"] = 302,
                ["03.354.923/0001-30"] = 303,
                ["22.191.947/0001-60"] = 304,
                ["58.068.537/0001-73"] = 305,
                ["05.058.384/0001-17"] = 306,
                ["52.170.102/0001-59"] = 307,
                ["27.348.657/0001-09"] = 308,
                ["71.929.830/0001-46"] = 309,
                ["04.402.050/0001-56"] = 310,
                ["08.266.773/0001-26"] = 311,
                ["22.062.969/0001-20"] = 312,
                ["04.969.068/0001-34"] = 313,
                ["14.997.950/0001-47"] = 314,
                ["18.972.142/0001-86"] = 315,
                ["02.558.157/0001-62"] = 318,
                ["23.604.315/0001-43"] = 319,
                ["08.798.784/0001-57"] = 320,
                ["03.933.877/0001-23"] = 321,
                ["26.803.992/0001-89"] = 322,
                ["29.019.953/0001-83"] = 323,
                ["10.736.174/0001-70"] = 324,
                ["40.588.835/0001-29"] = 325,
                ["09.646.089/0001-32"] = 326,
                ["66.970.229/0001-67"] = 329,
                ["35.147.498/0001-02"] = 330,
                ["06.159.860/0001-59"] = 340,
                ["08.326.545/0001-02"] = 345,
                ["09.149.703/0001-50"] = 436,
                ["02.558.157/0001-62"] = 451,
                ["61.695.227/0001-93"] = 459
            };
            
            if (cnpj != null && contratoFornecedores.TryGetValue(cnpj, out int idContrato))
            {
                return idContrato;
            }
            return 0;
        }
        
        public object ExtrairApenasNumeros(object obj, Type tipoDestino)
        {
            if (obj == null) 
                return tipoDestino == typeof(string) ? string.Empty : 0;

            string apenasDigitos = Regex.Replace(obj.ToString() ?? "", @"[^\d]", "");

            if (string.IsNullOrEmpty(apenasDigitos))
                return tipoDestino == typeof(string) ? string.Empty : 0;
            
            if (tipoDestino == typeof(string))
                return apenasDigitos;

            if (tipoDestino == typeof(int))
            {
                if (int.TryParse(apenasDigitos, out int resultado))
                {
                    return resultado;
                }
                return 0;
            }

            return apenasDigitos;
        }

    }       
        //contrato.NaturezaContratacao = (string) ExtrairApenasNumeros(valorExcel, typeof(string));

        //// Para o Pagamento (Tipo Int)
        //contrato.PagamentoParcelado = (int) ExtrairApenasNumeros(valorExcel, typeof(int));

}