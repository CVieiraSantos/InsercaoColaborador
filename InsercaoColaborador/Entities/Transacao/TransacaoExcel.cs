using InsercaoColaborador.Entities.Validation.Excecao;

namespace InsercaoColaborador.Entities.Transacao
{
    public class TransacaoExcel
    {
        public string NumeroCheque { get; set; } = null!;
        public DateTime DataPagamento { get; set; }
        public DateTime DataEmissaoDocFiscal { get; set; }
        public int EstadoEmissor { get; set; }
        public string ServicoProduto { get; set; } = null!;
        public decimal Total { get; set; }
        public string? NotaFiscalOuEquivalente { get; set; }
        public string NomeCredor { get; set; } = null!;
        public string CnpjCpf { get; set; } = null!;
        public decimal? ValorGlosado { get; set; }
        public string? ValorGlosadoString { get; set; }
        public bool ValorConciliado { get; set; }
        public string? Item { get; set; }
        public string? Observacoes { get; set; }
        public string? StatusAnalise { get; set; }
        public string? StatusCor { get; set; }
        public string Documento { get; set; } = null!;
        public string TipoDespesa { get; set; } = null!;
        public string SubCategoriaDeDespesa { get; set; } = null!;
        public string? Justificativa { get; set; }
        public string? ApuracaoGlosaParcial { get; set; }
        public decimal ValorBruto { get; set; }
        public decimal ValorEncargos { get; set; }
        public int? Rateio { get; set; }
        public decimal? PercentualRateio { get; set; }
        public string NumeroDoContrato { get; set; } = null!;
        //public Entities.Contrato.ContratoExcel ContratoExcel { get; set; }

        public static bool TryCreateTransacaoExcel
            (
                string item,
                string notaFiscalOuEquivalente,
                DateTime dataEmissaoDocFiscal,
                int estadoEmissor,
                string cnpjCpf,
                decimal ValorBruto,
                decimal valorEncargos,
                string SubCategoriaDeDespesa,
                int rateio,
                string numeroCheque,
                out TransacaoExcel transacaoExcel,
                List<string> errors
            )
        {
            if (string.IsNullOrWhiteSpace(item))
                throw new TransacaoExcelValidationException(
                    "O item da transação é obrigatória.",
                    nameof(item));
            if (string.IsNullOrWhiteSpace(notaFiscalOuEquivalente))
                throw new TransacaoExcelValidationException(
                    "A nota fiscal é obrigatória.",
                    nameof(notaFiscalOuEquivalente));

            if(dataEmissaoDocFiscal == default)
                throw new TransacaoExcelValidationException(
                    "A data da emissão fiscal é obrigatória.",
                    nameof(dataEmissaoDocFiscal));
            if(estadoEmissor <= 0)
                throw new TransacaoExcelValidationException(
                    "O valor do estado emissor precisa ser maior ou igual a 1.",
                    nameof(estadoEmissor));
            if (string.IsNullOrWhiteSpace(cnpjCpf))
                throw new TransacaoExcelValidationException(
                    "A nota fiscal é obrigatória.",
                    nameof(cnpjCpf));
            if(ValorBruto <= 0.0m)
                throw new TransacaoExcelValidationException(
                    "O valor bruto precisa ser maior ou igual a 1.",
                    nameof(ValorBruto));
            if (valorEncargos <= 0.0m)
                throw new TransacaoExcelValidationException(
                    "O valor encargos precisa ser maior ou igual a 1.",
                    nameof(valorEncargos));
            if (string.IsNullOrWhiteSpace(SubCategoriaDeDespesa))
                throw new TransacaoExcelValidationException(
                    "A categoria de despesa é obrigatória.",
                    nameof(SubCategoriaDeDespesa));
            if(rateio < 0)
                throw new TransacaoExcelValidationException(
                    "O rateio é obrigatório podendo ser zero, mais nunca vazio ou menor que zero.",
                    nameof(rateio));
            if (string.IsNullOrWhiteSpace(numeroCheque))
                throw new TransacaoExcelValidationException(
                    "O número do contrato é obrigatório.",
                    nameof(numeroCheque));

            if (errors.Any())
            {
                transacaoExcel = null!;
                return false;
            }

            transacaoExcel = new TransacaoExcel
            {
                Item = item.Trim(),
                NotaFiscalOuEquivalente = notaFiscalOuEquivalente.Trim(),
                DataEmissaoDocFiscal = dataEmissaoDocFiscal,
                EstadoEmissor = estadoEmissor,
                CnpjCpf = cnpjCpf.Trim(),
                ValorBruto = ValorBruto,
                ValorEncargos = valorEncargos,
                SubCategoriaDeDespesa = SubCategoriaDeDespesa.Trim(),
                Rateio = rateio,
                NumeroCheque = numeroCheque.Trim(),
            };
            
            return true;
        }


    }
}
