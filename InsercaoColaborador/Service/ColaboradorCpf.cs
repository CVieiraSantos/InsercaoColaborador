using InsercaoColaborador.Entities.Colaborador;

namespace InsercaoColaborador.Service
{
    public static class ColaboradorCpf
    {
        public static List<Colaborador> FiltroCpfUnico(IEnumerable<Colaborador> colaboradors)
        {
            var employee = colaboradors
                .Where(c => !string.IsNullOrWhiteSpace(CpfCnpjGenerator.FormatarCpf(c.CPF)))
                .DistinctBy(c => c.CPF)
                .Select(x => new Colaborador
                {
                    Nome = x.Nome,
                    CPF = x.CPF,
                    CNS = x.CNS
                })
                .ToList();

            return employee;
        }


    }
}
