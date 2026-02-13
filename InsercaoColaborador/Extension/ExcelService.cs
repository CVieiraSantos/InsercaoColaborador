using ClosedXML.Excel;

namespace InsercaoColaborador.Extension
{
    public class ExcelService
    {
        public static List<T> ImportarExcel<T>(string caminhoExcel, object identificadorAba, Func<IXLRangeRow, T?> mapeador)
        {
            var lista = new List<T>();

            using (var fs = File.Open(caminhoExcel, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var workbook = new XLWorkbook(fs))
            {
                IXLWorksheet planilha;
                if (identificadorAba is int indice)
                    planilha = workbook.Worksheet(indice);
                else
                    planilha = workbook.Worksheet(identificadorAba.ToString());

                var linhas = planilha.RangeUsed()?.RowsUsed()?.Skip(1) ?? Enumerable.Empty<IXLRangeRow>();

                foreach (var linha in linhas)
                {
                    var item = mapeador(linha);
                    if (item != null) lista.Add(item);
                }
            }
            return lista;
        }


    }
}
