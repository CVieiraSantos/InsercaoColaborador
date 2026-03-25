using ClosedXML.Excel;
using System.Globalization;

namespace InsercaoColaborador.Extension
{
    public static class ValorEmDecimal
    {
        public static decimal GetDecimal(this IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return 0m;

            var formatted = cell.GetFormattedString()?.Trim();

            if (string.IsNullOrWhiteSpace(formatted) || formatted == "-" || formatted == "N/A")
                return 0m;

            if (formatted == "1%")
                return 1.0m;

            if (formatted.Contains("%"))
            {
                var text = formatted.Replace("%", "").Trim();

                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out var percent))
                {
                    return percent;
                }
            }

            return cell.GetValue<decimal>();
        }
    }
}
