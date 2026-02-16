using System.Text;

namespace InsercaoColaborador.Infrastructure.Sql.Builders
{
    public static class SqlInsertBuilder
    {
        public static string BuildInsert<T>(
        string table,
        IReadOnlyList<string> columns,
        IEnumerable<T> items,
        Func<T, string> valuesProjection)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"INSERT INTO {table}");
            sb.AppendLine("(");
            sb.AppendLine("    " + string.Join(", ", columns));
            sb.AppendLine(") VALUES");

            var tuples = items.Select(valuesProjection);

            sb.AppendLine(string.Join("," + Environment.NewLine, tuples));
            sb.Append(";");

            return sb.ToString();
        }
    }
}
