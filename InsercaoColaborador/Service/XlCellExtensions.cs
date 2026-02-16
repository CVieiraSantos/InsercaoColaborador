using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsercaoColaborador.Service
{
    public static class XlCellExtensions
    {
        public static DateTime? GetDateTimeOrNull(this IXLCell cell, CultureInfo? culture = null)
        {
            if (cell.Value.IsDateTime)
                return cell.GetDateTime();

            var s = cell.GetString()?.Trim();

            if (string.IsNullOrEmpty(s))
                return null;

            if (culture == null)
                culture = new CultureInfo("pt-BR");

            if (DateTime.TryParse(s, culture, DateTimeStyles.None, out var dt))
                return dt;

            return null;
        }
    }
}
