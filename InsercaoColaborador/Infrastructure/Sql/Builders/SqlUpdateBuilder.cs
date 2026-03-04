using System.Text;

namespace InsercaoColaborador.Infrastructure.Sql.Builders
{
    public static class SqlUpdateBuilder
    {
        public static string BuildUpdates<T>(
            string table,
            IEnumerable<T> items,
            Func<T, string> setProjection,
            Func<T, string> whereProjection)
        {
            if (items is null) throw new ArgumentNullException(nameof(items));
            if (setProjection is null) throw new ArgumentNullException(nameof(setProjection));
            if (whereProjection is null) throw new ArgumentNullException(nameof(whereProjection));

            var list = items.ToList();
            if (!list.Any()) return string.Empty;

            var sb = new StringBuilder();

            foreach (var item in list)
            {
                var setClause = setProjection(item) ?? throw new InvalidOperationException("setProjection returned null");
                var whereClause = whereProjection(item) ?? throw new InvalidOperationException("whereProjection returned null");

                sb.AppendLine($"UPDATE {table}");
                sb.AppendLine("SET");
                sb.AppendLine("    " + setClause);
                sb.AppendLine("WHERE " + whereClause + ";");
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
