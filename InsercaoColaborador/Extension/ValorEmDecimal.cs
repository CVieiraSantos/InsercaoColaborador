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

            if (cell.DataType == XLDataType.Text)
            {
                var text = cell.GetString();
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
                    return result;
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                    return result;
            }

            try
            {
                return cell.GetValue<decimal>();
            }
            catch
            {
                var text = cell.GetString();
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out var result))
                    return result;
                if (decimal.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                    return result;

                throw;
            }
        }
    }
}
