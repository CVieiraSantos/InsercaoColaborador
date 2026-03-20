using ClosedXML.Excel;
using System.Globalization;


namespace InsercaoColaborador.Extension
{
    public static class ValorEmData
    {
        private static readonly string[] NullTokens = new[] { "-", "N/A", "NA", "SEM DATA", "S/D", "NULL" };
        public static DateTime? GetDatetime(this IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return null;

            if (cell.DataType == XLDataType.Number)
            {
                try
                {
                    var dnum = cell.GetValue<double>();
                    return DateTime.FromOADate(dnum);
                }
                catch {}
            }

            var s = cell.GetString()?.Trim();
            if (string.IsNullOrEmpty(s)) return null;

            if (NullTokens.Contains(s, StringComparer.OrdinalIgnoreCase)) return DateTime.MinValue;

            if (DateTime.TryParse(s, new CultureInfo("pt-BR"), DateTimeStyles.None, out DateTime dta))
                return dta;

            var formats = new[] { "M/d/yyyy", "M/d/yy", "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };
            if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                return dt;
            if (DateTime.TryParse(s, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                return dt;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                return dt;

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                return DateTime.FromOADate(d);

            return null;
        }
    }
}
