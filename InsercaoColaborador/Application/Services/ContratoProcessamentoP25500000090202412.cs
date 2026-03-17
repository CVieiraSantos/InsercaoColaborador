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
        public string Nome => "Gerar INSERT de Contratos Modelo Importacao Contratos P25500000090202412";

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
                    NomeBeneficiario = linha.Cell(1).GetString().Trim(),
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
                    Objeto = linha.Cell(15).GetString(),
                    NaturezaDeContratacao = linha.Cell(16).GetString().Trim(),
                    NaturezaNaoEspecificada = linha.Cell(17).GetString().Trim(),
                    ArtigoRegulamentoCompras = linha.Cell(18).GetString().Trim()
                };
            });

            var contratos = contratoExcel.Select(e =>
            {
                return new Contrato
                {
                    IdFornecedor = ObterIdPorCnpjPorFornecedor(e.CnpjFornecedor),
                    RazaoSocialFornecedor = e.NomeBeneficiario?.Trim().Length > 100
                    ? e.NomeBeneficiario.Substring(0,100) 
                    : e.NomeBeneficiario ?? string.Empty,
                    NumeroContrato = e.NumeroContrato.Trim(),
                    Parcelado = (int)ExtrairApenasNumeros(e.PagamentoParcelado, typeof(int)),
                    QuantidadedeParcelas = e.QuantidadedeParcelas,
                    TipoValorContrato = (int)ExtrairApenasNumeros(e.TipoDeValorDoContrato,typeof(int)),
                    TipoVigencia = (int)ExtrairApenasNumeros(e.TipoDeVigencia, typeof(int)),
                    Inicio = e.InicioVigencia,
                    Fim = e.FimVigencia,
                    DataAssinatura = e.DataAssinatura,
                    CriterioSelecao = (int)ExtrairApenasNumeros(e.CriterioDeSelecao, typeof(int)),
                    CriterioSelecaoOutro = e.CriterioDeSelecao == 4 ? e.CriterioDeSelecaoOutro : string.Empty,
                    CategoriaDespesa = e.CategoriaDeDespesa.Trim(),
                    Valor = e.Valor,
                    ObjetoContrato = e.Objeto,
                    NaturezaContratacao = (string)ExtrairApenasNumeros(e.NaturezaDeContratacao,typeof(string)),
                    NaturezaContratacaoOutro = e.NaturezaDeContratacao == "23" ? e.NaturezaNaoEspecificada : string.Empty,
                    ArtigoRegulamentoCompras = e.ArtigoRegulamentoCompras,
                    TipoFornecedor = 1,
                    IdIdentidade = 75,
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

        public int ObterIdPorCnpjPorFornecedor(string cnpj)
        {
            if (string.IsNullOrEmpty(cnpj)) return 0;

            var fornecedores = new List<KeyValuePair<string, int>>
            {
                new KeyValuePair<string, int>("48.549.331/0001-01", 287),
                new KeyValuePair<string, int>("02.287.620/0001-89", 288),
                new KeyValuePair<string, int>("44.166.627/0001-92", 289),
                new KeyValuePair<string, int>("03.858.331/0001-55", 290),
                new KeyValuePair<string, int>("23.344.957/0001-50", 291),
                new KeyValuePair<string, int>("04.436.106/0001-93", 292),
                new KeyValuePair<string, int>("61.695.227/0001-93", 293),
                new KeyValuePair<string, int>("08.306.796/0001-17", 294),
                new KeyValuePair<string, int>("23.236.675/0001-30", 295),
                new KeyValuePair<string, int>("22.314.093/0001-61", 296),
                new KeyValuePair<string, int>("32.643.177/0001-00", 297),
                new KeyValuePair<string, int>("23.035.684/0001-62", 298),
                new KeyValuePair<string, int>("45.728.775/0001-16", 299),
                new KeyValuePair<string, int>("09.294.688/0001-34", 300),
                new KeyValuePair<string, int>("29.016.530/0001-00", 301),
                new KeyValuePair<string, int>("03.252.307/0001-78", 302),
                new KeyValuePair<string, int>("03.354.923/0001-30", 303),
                new KeyValuePair<string, int>("22.191.947/0001-60", 304),
                new KeyValuePair<string, int>("58.068.537/0001-73", 305),
                new KeyValuePair<string, int>("05.058.384/0001-17", 306),
                new KeyValuePair<string, int>("52.170.102/0001-59", 307),
                new KeyValuePair<string, int>("27.348.657/0001-09", 308),
                new KeyValuePair<string, int>("71.929.830/0001-46", 309),
                new KeyValuePair<string, int>("04.402.050/0001-56", 310),
                new KeyValuePair<string, int>("08.266.773/0001-26", 311),
                new KeyValuePair<string, int>("22.062.969/0001-20", 312),
                new KeyValuePair<string, int>("04.969.068/0001-34", 313),
                new KeyValuePair<string, int>("14.997.950/0001-47", 314),
                new KeyValuePair<string, int>("18.972.142/0001-86", 315),
                new KeyValuePair<string, int>("02.558.157/0001-62", 318),
                new KeyValuePair<string, int>("23.604.315/0001-43", 319),
                new KeyValuePair<string, int>("08.798.784/0001-57", 320),
                new KeyValuePair<string, int>("03.933.877/0001-23", 321),
                new KeyValuePair<string, int>("26.803.992/0001-89", 322),
                new KeyValuePair<string, int>("29.019.953/0001-83", 323),
                new KeyValuePair<string, int>("10.736.174/0001-70", 324),
                new KeyValuePair<string, int>("40.588.835/0001-29", 325),
                new KeyValuePair<string, int>("09.646.089/0001-32", 326),
                new KeyValuePair<string, int>("66.970.229/0001-67", 329),
                new KeyValuePair<string, int>("35.147.498/0001-02", 330),
                new KeyValuePair<string, int>("06.159.860/0001-59", 340),
                new KeyValuePair<string, int>("08.326.545/0001-02", 345),
                new KeyValuePair<string, int>("09.149.703/0001-50", 436),
                new KeyValuePair<string, int>("02.558.157/0001-62", 451),
                new KeyValuePair<string, int>("61.695.227/0001-93", 459)
            };

            var resultado = fornecedores.FirstOrDefault(x => x.Key == cnpj);

            return resultado.Value;
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
}