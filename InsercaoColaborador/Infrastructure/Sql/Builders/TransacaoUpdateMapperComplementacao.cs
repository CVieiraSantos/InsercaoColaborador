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
            sb.Append($"NotaFiscal = {item.NotaFiscal.ToSql()},\n");
            sb.Append($"DataNotaFiscal = {item.DataNotaFiscal.ToSql()},\n");
            sb.Append($"EstadoEmissor = {item.EstadoEmissor.ToSql()},\n");
            sb.Append($"NomeBeneficiario = {item.NomeBeneficiario.ToSql()},\n");
            sb.Append($"Valor = {item.Valor.ToSql()},\n");
            sb.Append($"ValorEncargos = {item.ValorEncargos.ToSql()},\n");
            sb.Append($"Categoria = {item.Categoria.ToSql()},\n");
            sb.Append($"ExisteRateio = {item.ExisteRateio.ToSql()},\n");
            sb.Append($"PercentualRateio = {item.PercentualRateio.ToSql()},\n");
            sb.Append($"NumeroDoContrato = {item.NumeroDoContrato.ToSql()}");
            return sb.ToString();
        }

        // WHERE por IdExtrato (recomendado se IdExtrato estiver preenchido)
        public static string MapWhereByIdExtrato(Transacao item) =>
            $"IdExtrato = {item.IdExtrato.ToSql()}";

        // Exemplo de WHERE por chave de negócio alternativa
        public static string MapWhereByAlternativeColumns(Transacao item) =>
            $"ObservacoesEntidade = {item.ObservacoesEntidade.ToSql()} " +
            $"AND IdCliente = {item.IdCliente.ToSql()} " +
            $"AND IdParceria = {item.IdParceria.ToSql()}";
    }
}
