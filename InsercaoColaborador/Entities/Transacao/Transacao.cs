namespace InsercaoColaborador.Entities.Transacao
{
    public class Transacao
    {
        public int IdExtrato { get; set; }
        public string Numero { get; set; } = null!;
        public DateTime DataTransacao { get; set; }
        public string Descricao { get; set; } = null!;
        public decimal Valor { get; set; }
        public string Tipo { get; set; } = null!;
        public string? NotaFiscal { get; set; }
        public string? Categoria { get; set; }
        public DateTime DataNotaFiscal { get; set; }
        public string NomeBeneficiario { get; set; } = null!;
        public string? OrigemRecurso { get; set; }
        public int IdParceria { get; set; }
        public int Referencia { get; set; }
        public int Exercicio { get; set; }
        public int Status { get; set; }
        public string? Avaliador { get; set; }
        public decimal? ValorContestado { get; set; }
        public int Conciliado { get; set; }
        public string? ObservacoesEntidade { get; set; }
        public string? ObservacoesOrgao { get; set; }
        public DateTime? DataHoraAnalise { get; set; }
        public string? AnaliseEscrita { get; set; }
        public int? IdRepasse { get; set; }
        public int? IdBeneficiario { get; set; }
        public string? TipoBeneficiario { get; set; }
        public DateTime? DataHoraCadastro { get; set; }
        public DateTime? DataHoraUltimaAlteracao { get; set; }
        public int IdCliente { get; set; }
        public int? NaturezaDevolucao { get; set; }
        public int? IdEBanco { get; set; }
        public int MeioPagamento { get; set; }
        public decimal ValorDocumento { get; set; }
        public decimal ValorEncargos { get; set; }
        public int EstadoEmissor { get; set; }
        public int? IdContrato { get; set; }
        public int SubCategoria { get; set; }
        public string? ItemDespesa { get; set; }
        public int? IdItemPlanoAplicacao { get; set; }
        public int ExisteRateio { get; set; }
        public decimal? PercentualRateio { get; set; }
        public DateTime? DataHoraConciliacao { get; set; }
        public int? IdUnidadeAtendimento { get; set; }
        public string? IdentificadorStorage { get; set; }
        public string? UrlDownloadArquivoTransacao { get; set; }     
    }
}
