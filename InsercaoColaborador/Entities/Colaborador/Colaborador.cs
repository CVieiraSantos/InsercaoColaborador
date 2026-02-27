namespace InsercaoColaborador.Entities.Colaborador
{
    public class Colaborador
    {
        public int IdTResponsavel { get; set; }
        public int IdCargo { get; set; }
        public int IdEntidade { get; set; }
        public string Nome { get; set; } = string.Empty;
        public DateTime DataNascimento { get; set; }
        public string Rg { get; set; } = null!;
        public string CPF { get; set; } = string.Empty;
        public string Endereco { get; set; } = null!;
        public string Numero { get; set; } = null!;
        public string Bairro { get; set; } = null!;
        public string Cidade { get; set; } = null!;
        public string Uf { get; set; } = null!;
        public string Cep { get; set; } = null!;
        public string? TelContato1 { get; set; }
        public string? TelContato2 { get; set; }
        public string? Email1 { get; set; }
        public string? Email2 { get; set; }
        public DateTime DataCriacao { get; set; }
        public int Ativo { get; set; }
        public string? OrgaoClasse { get; set; }
        public string Formacao { get; set; } = null!;
        public string Vinculo { get; set; } = null!;
        public int CargaHoraria { get; set; }
        public decimal Salario { get; set; }
        public int IdCliente { get; set; }
        public string? CNS { get; set; }
    }
}
