using InsercaoColaborador.Entities.Colaborador;
using System.Text.RegularExpressions;

namespace InsercaoColaborador.Service
{
    public static class CpfCnpjGenerator
    {
        public static string FormatarCpfOuCnpj(string valor)
        {
            if (string.IsNullOrWhiteSpace(valor))
                return "";

            var numeros = Regex.Replace(valor, @"\D", "");

            if (numeros.Length == 11)
                return Regex.Replace(numeros, @"(\d{3})(\d{3})(\d{3})(\d{2})", "$1.$2.$3-$4");

            if (numeros.Length == 14)
                return Regex.Replace(numeros, @"(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})", "$1.$2.$3/$4-$5");

            return numeros;
        }

        public static bool EhCpf(string valor)
            => Regex.Replace(valor ?? "", @"\D", "").Length == 11;

        public static bool EhCnpj(string valor)
            => Regex.Replace(valor ?? "", @"\D", "").Length == 14;

        public static string FormatarCpf(string? cpf)
        {
            if (string.IsNullOrWhiteSpace(cpf))
                return "";

            var numeros = Regex.Replace(cpf, @"\D", "");

            if (numeros.Length == 11)
            {
                return Regex.Replace(numeros, @"(\d{3})(\d{3})(\d{3})(\d{2})", "$1.$2.$3-$4");
            }

            return "";
        }

        public static string FormatarCpf(Colaborador colaboradorExcel)
        {
            if (colaboradorExcel is null)
                return "";

            return FormatarCpf(colaboradorExcel.CPF);
        }

    }
}