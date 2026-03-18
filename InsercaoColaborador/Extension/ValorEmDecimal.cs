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

            if (cell.DataType == XLDataType.Number)
                return cell.GetValue<decimal>();

            var text = cell.GetString().Trim();

            if (text == "-" || text == "N/A" || string.IsNullOrWhiteSpace(text))
                return 0m;

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
                return result;

            if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                return result;

            try
            {
                return cell.GetValue<decimal>();
            }
            catch
            {
                return 0m;
            }
        }
    }
}
