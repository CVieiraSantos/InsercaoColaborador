using InsercaoColaborador.Entities.Transacao;

namespace InsercaoColaborador.Service
{
    public static class GerarNomeBeneficiario
    {
        public static string GetNomeBeneficiario(TransacaoExcel e)
        {
            if (!string.IsNullOrWhiteSpace(e.NomeCredor) && e.NomeCredor.Contains('-'))
                return e.NomeCredor;

            return e.CnpjCpf;
        }
    }
}
