using System.Globalization;

namespace InsercaoColaborador.Infrastructure.Sql.ConverterSql
{
    public static class SqlValueExtensions
    {
        public static string ToSql(this string? s) =>
            s == null ? "NULL" : $"'{s.Replace("'", "''")}'";

        public static string ToSql(this int i) =>
            i.ToString(CultureInfo.InvariantCulture);

        public static string ToSql(this int? i) =>
            i.HasValue ? i.Value.ToSql() : "NULL";

        public static string ToSql(this decimal d) =>
            d.ToString("F2", CultureInfo.InvariantCulture);
        
        public static string ToSql(this double d) =>
            d.ToString("F2", CultureInfo.InvariantCulture);

        public static string ToSql(this decimal? d) =>
            d.HasValue ? d.Value.ToSql() : "NULL";

        public static string ToSql(this DateTime dt) =>
            $"'{dt:yyyy-MM-dd HH:mm:ss}'";

        public static string ToSql(this DateTime? dt) =>
            dt.HasValue ? dt.Value.ToSql() : "NULL";
    }
}
