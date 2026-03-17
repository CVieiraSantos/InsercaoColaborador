namespace InsercaoColaborador.Entities.Contrato
{
    public class ContratoExcel
    {
        public int IdFornecedor { get; set; }
        public string CnpjFornecedor { get; set; } = null!;
        public string? NomeBeneficiario { get; set; }
        public string NumeroContrato { get; set; } = null!;
        public int PagamentoParcelado { get; set; }
        public int QuantidaParcelas { get; set; }
        public int TipoDeValorDoContrato { get; set; }
        public int TipoDeVigencia { get; set; }
        public DateTime InicioVigencia { get; set; }
        public DateTime FimVigencia { get; set; }
        public DateTime? DataAssinatura { get; set; }
        public int CriterioDeSelecao { get; set; }
        public string? CriterioDeSelecaoOutro { get; set; }
        public string CategoriaDeDespesa { get; set; } = null!;
        public double Valor { get; set; }
        public string Objeto { get; set; } = null!;
        public string NaturezaDeContratacao { get; set; } = null!;
        public string? NaturezaNaoEspecificada { get; set; }
        public string? ArtigoRegulamentoCompras { get; set; }

    }
}
