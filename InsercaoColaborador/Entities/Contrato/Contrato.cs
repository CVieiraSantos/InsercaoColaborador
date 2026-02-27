namespace InsercaoColaborador.Entities.Contrato
{
    public class Contrato
    {
        public int IdContrato { get; set; }
        public string NumeroContrato { get; set; } = null!;
        public int? IdFornecedor { get; set; }
        public int? IdCotacao { get; set; }
        public string RazaoSocialFornecedor { get; set; } = null!;
        public double Valor { get; set; }
        public DateTime Inicio { get; set; }
        public DateTime Fim { get; set; }
        public int Parcelado { get; set; }
        public int QuantidadedeParcelas { get; set; }
        public string CategoriaDespesa { get; set; } = null!;
        public string? LinkArquivoContrato { get; set; }
        public int IdIdentidade { get; set; }
        public int Ativo { get; set; }
        public int IdCliente { get; set; }
        public int? IdPrestadorServico { get; set; }
        public int TipoValorContrato { get; set; }
        public int TipoVigencia { get; set; }
        public DateTime? DataAssinatura { get; set; }
        public int CriterioSelecao { get; set; }
        public string ObjetoContrato { get; set; } = null!;
        public string? CriterioSelecaoOutro { get; set; }
        public string? ArtigoRegulamentoCompras { get; set; }
        public string NaturezaContratacao { get; set; } = null!;
        public int TipoFornecedor { get; set; }
        public string? NaturezaContratacaoOutro { get; set; }
    }
}
