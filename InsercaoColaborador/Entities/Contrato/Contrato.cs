using InsercaoColaborador.Entities.Validation.Excecao;

namespace InsercaoColaborador.Entities.Contrato
{
    public class Contrato
    {
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

        public static bool TryCreate(
        string razaoSocialFornecedor,
        string numeroContrato,
        int parcelado,
        int quantidadeDeParcelas,
        int tipoValorContrato,
        int tipoVigencia,
        DateTime inicio,
        DateTime fim,
        DateTime dataAssinatura,
        int criterioSelecao,
        string criterioSelecaoOutro,
        string categoriaDespesa,
        double valor,
        string objetoContrato,
        string naturezaContratacao,
        out Contrato contrato,
        out List<string> errors)
        {
            errors = new List<string>();

            // Strings obrigatórias → exceção
            if (string.IsNullOrWhiteSpace(razaoSocialFornecedor))
                throw new ContratoValidationException(
                    "Razão social do fornecedor é obrigatória.",
                    nameof(razaoSocialFornecedor));

            if (string.IsNullOrWhiteSpace(numeroContrato))
                throw new ContratoValidationException(
                    "Número do contrato é obrigatório.",
                    nameof(numeroContrato));

            if (string.IsNullOrWhiteSpace(categoriaDespesa))
                throw new ContratoValidationException(
                    "Categoria de despesa é obrigatória.",
                    nameof(categoriaDespesa));

            if (string.IsNullOrWhiteSpace(objetoContrato))
                throw new ContratoValidationException(
                    "Objeto do contrato é obrigatório.",
                    nameof(objetoContrato));

            if (string.IsNullOrWhiteSpace(naturezaContratacao))
                throw new ContratoValidationException(
                    "Natureza da contratação é obrigatória.",
                    nameof(naturezaContratacao));

            // Validações de regras de negócio (sem exceção)
            if (parcelado is not(0 or 1))
                throw new ContratoValidationException(
                    "Pagamento parcelado é obrigatório e precisa ser 0 ou 1.",
                    nameof(parcelado));

            if (valor <= 0)
                throw new ContratoValidationException(
                    "O valor é obrigatório e precisa ser maior que zero.",
                    nameof(valor));

            if (fim == default)
            throw new ContratoValidationException(
                "O campo Fim é obrigatório.",
                nameof(fim));

            if(criterioSelecao is not(1 or 2 or 3 or 4))
                throw new ContratoValidationException(
                "O campo Critério de seleção não pode ser nulo ou vazio, tampoco ser diferente de 1,2,3,4",
                nameof(criterioSelecao));

            if (errors.Any())
            {
                contrato = null!;
                return false;
            }

            contrato = new Contrato
            {
                NumeroContrato = numeroContrato.Trim(),
                RazaoSocialFornecedor = razaoSocialFornecedor.Trim(),
                CategoriaDespesa = categoriaDespesa.Trim(),
                ObjetoContrato = objetoContrato.Trim(),
                NaturezaContratacao = naturezaContratacao.Trim(),
                Parcelado = parcelado,
                QuantidadedeParcelas = quantidadeDeParcelas,
                TipoValorContrato = tipoValorContrato,
                TipoVigencia = tipoVigencia,
                Inicio = inicio,
                Fim = fim,
                DataAssinatura = dataAssinatura,
                CriterioSelecao = criterioSelecao,
                CriterioSelecaoOutro = criterioSelecaoOutro?.Trim(),
                Valor = valor
            };

            return true;
        }


    }
}
