namespace InsercaoColaborador.Entities.Transacao
{
    public class TransacaoExcel
    {
        public string NumeroCheque { get; set; } = null!;
        public DateTime DataPagamento { get; set; }
        public string ServicoProduto { get; set; } = null!;
        public decimal Total { get; set; }
        public string? NotaFiscalOuEquivalente { get; set; }
        public string NomeCredor { get; set; } = null!;
        public string CnpjCpf { get; set; } = null!;
        public decimal? ValorGlosado { get; set; }
        public bool ValorConciliado { get; set; }
        public string? Item { get; set; }
        public string? Observacoes { get; set; }
        public string StatusAnalise { get; set; }
        public string? StatusCor { get; set; }
        public string Documento { get; set; } = null!;
    }
}
