using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Infrastructure.Sql.ConverterSql;
using System.Text;

namespace InsercaoColaborador.Infrastructure.Sql.Builders
{
    public static class TransacaoUpdateMapperComplementacao
    {
        public static string MapUpdateSet(Transacao item)
        {
            var sb = new StringBuilder();
            if(item.DataNotaFiscal != DateTime.MinValue)
                sb.Append($"DataNotaFiscal = {item.DataNotaFiscal.ToSql()},\n");
            sb.Append($"EstadoEmissor = {item.EstadoEmissor.ToSql()},\n");
            sb.Append($"ValorDocumento = {item.ValorDocumento.ToSql()},\n");
            sb.Append($"ValorEncargos = {item.ValorEncargos.ToSql()},\n");
            sb.Append($"Categoria = {item.Categoria.ToSql()},\n");
            sb.Append($"ExisteRateio = {item.ExisteRateio.ToSql()},\n");
            sb.Append($"PercentualRateio = {item.PercentualRateio.ToSql()}\n");
            if(item.IdContrato > 0 && item.IdContrato != null)
                sb.Append($", IdContrato = {item.IdContrato.ToSql()}");
            return sb.ToString();
        }

        // WHERE por IdExtrato (recomendado se IdExtrato estiver preenchido)
        public static string MapWhereByIdExtrato(Transacao item) =>
            $"IdExtrato = {item.IdExtrato.ToSql()}";

        // Exemplo de WHERE por chave de negócio alternativa
        public static string MapWhereByAlternativeColumns(Transacao item) =>
            $"ObservacoesEntidade = {item.ObservacoesEntidade.ToSql()} AND " +
            $"IdCliente = {item.IdCliente.ToSql()} AND " +
            $"IdParceria = {item.IdParceria.ToSql()}";
    }
}
