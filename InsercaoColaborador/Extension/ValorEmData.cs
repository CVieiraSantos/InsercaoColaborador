using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsercaoColaborador.Extension
{
    public static class ValorEmData
    {
        public static DateTime? GetDatetime(this IXLCell cell)
        {

            var s = cell.GetString()?.Trim();
            if (string.IsNullOrEmpty(s)) return null;

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
