using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Infrastructure.Sql.ConverterSql;

namespace InsercaoColaborador.Infrastructure.Sql.Mappers.TransacaoMapper
{
    public class TransacaoSqlMapper
    {
        public static string MapValues(Transacao item) => $@"(
        {item.IdExtrato.ToSql()},
        {item.Numero.ToSql()},
        {item.DataTransacao.ToSql()},
        {item.Descricao.ToSql()},
        {item.Valor.ToSql()},
        {item.Tipo.ToSql()},
        {item.NotaFiscal.ToSql()},
        {item.Categoria.ToSql()},
        {item.DataNotaFiscal.ToSql()},
        {item.NomeBeneficiario.ToSql()},
        {item.OrigemRecurso.ToSql()},
        {item.IdParceria.ToSql()},
        {item.Referencia.ToSql()},
        {item.Exercicio.ToSql()},
        {item.Status.ToSql()},
        {item.Avaliador.ToSql()},
        {item.ValorContestado.ToSql()},
        {item.Conciliado.ToSql()},
        {item.DataHoraConciliacao.ToSql()},
        {item.ObservacoesEntidade.ToSql()},
        {item.ObservacoesOrgao.ToSql()},
        {item.DataHoraCadastro.ToSql()},
        {item.IdCliente.ToSql()},
        {item.IdEBanco.ToSql()},
        {item.MeioPagamento.ToSql()},
        {item.ValorDocumento.ToSql()},
        {item.ValorEncargos.ToSql()},
        {item.EstadoEmissor.ToSql()},
        {item.SubCategoria.ToSql()},
        {item.ExisteRateio.ToSql()},
        {item.AnaliseEscrita.ToSql()}
    )";
    }
}
