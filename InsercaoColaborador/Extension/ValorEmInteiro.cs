using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InsercaoColaborador.Extension
{
    public static class ValorEmInteiro
    {
        public static int GetInt(this IXLCell cell)
        {
            if (cell.TryGetValue<int>(out var valor))
                return valor;

            var texto = cell.GetString().Trim();

            if (int.TryParse(texto, out valor))
                return valor;

            throw new FormatException("Valor inválido");
        }
    }
}
