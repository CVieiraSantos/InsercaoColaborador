using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace InsercaoColaborador.Extension
{
    public static class ValorEmInteiro
    {
        public static int GetInt(this IXLCell cell)
        {
            if (cell.TryGetValue<int>(out var valor))
                return valor;

            var texto = cell.GetString().Trim();

            if (string.IsNullOrEmpty(texto))
                return 0;

            string apenasDigitos = Regex.Replace(texto, @"[^\d]", "");

            if (int.TryParse(apenasDigitos, out int resultado))
            {
                return resultado;
            }

            return 0;
        }
    }
}
