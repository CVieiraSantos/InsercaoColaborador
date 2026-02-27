using InsercaoColaborador.Entities.Colaborador;
using InsercaoColaborador.Infrastructure.Sql.ConverterSql;

namespace InsercaoColaborador.Infrastructure.Sql.Mappers.ColaboradorMapper
{
    public class ColaboradorSqlMapper
    {
        public static string MapValues(Colaborador item) => $@"(
            {item.IdTResponsavel.ToSql()},
            {item.IdCargo.ToSql()},
            {item.IdEntidade.ToSql()},
            {item.Nome.ToSql()},
            {item.DataNascimento.ToSql()},
            {item.Rg.ToSql()},
            {item.CPF.ToSql()},
            {item.Endereco.ToSql()},
            {item.Numero.ToSql()},
            {item.Bairro.ToSql()},
            {item.Cidade.ToSql()},
            {item.Uf.ToSql()},
            {item.Cep.ToSql()},
            {item.TelContato1.ToSql()},
            {item.TelContato2.ToSql()},
            {item.Email1.ToSql()},
            {item.Email2.ToSql()},
            {item.DataCriacao.ToSql()},
            {item.Ativo.ToSql()},
            {item.OrgaoClasse.ToSql()},
            {item.Formacao.ToSql()},
            {item.Vinculo.ToSql()},
            {item.CargaHoraria.ToSql()},
            {item.Salario.ToSql()},
            {item.IdCliente.ToSql()},
            {item.CNS.ToSql()}
        )";
    }
}
