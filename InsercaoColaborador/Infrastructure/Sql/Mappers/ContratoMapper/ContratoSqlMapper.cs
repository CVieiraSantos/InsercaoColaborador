using InsercaoColaborador.Entities.Contrato;
using InsercaoColaborador.Infrastructure.Sql.ConverterSql;

namespace InsercaoColaborador.Infrastructure.Sql.Mappers.ContratoMapper
{
    public class ContratoSqlMapper
    {
        public static string MapValues(Contrato item) => $@"(
            {item.IdContrato.ToSql()},
            {item.NumeroContrato.ToSql()},
            {item.IdFornecedor.ToSql()},
            {item.IdCotacao.ToSql()},
            {item.RazaoSocialFornecedor.ToSql()},
            {item.Valor.ToSql()},
            {item.Inicio.ToSql()},
            {item.Fim.ToSql()},
            {item.Parcelado.ToSql()},
            {item.QuantidadedeParcelas.ToSql()},
            {item.CategoriaDespesa.ToSql()},
            {item.LinkArquivoContrato.ToSql()},
            {item.IdIdentidade.ToSql()},
            {item.Ativo.ToSql()},
            {item.IdCliente.ToSql()},
            {item.IdPrestadorServico.ToSql()},
            {item.TipoValorContrato.ToSql()},
            {item.TipoVigencia.ToSql()},
            {item.DataAssinatura.ToSql()},
            {item.CriterioSelecao.ToSql()},
            {item.ObjetoContrato.ToSql()},
            {item.CriterioSelecaoOutro.ToSql()},
            {item.ArtigoRegulamentoCompras.ToSql()},
            {item.NaturezaContratacao.ToSql()},
            {item.TipoFornecedor.ToSql()},
            {item.NaturezaContratacaoOutro.ToSql()}
        )";
    }
}
