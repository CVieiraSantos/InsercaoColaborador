
namespace InsercaoColaborador.Entities.Validation.Excecao
{
    public class TransacaoExcelValidationException : Exception
    {
        public string FieldName { get; }

        public TransacaoExcelValidationException(string message, string fieldName) : base(message)
        {
            FieldName = fieldName;
        }

        public TransacaoExcelValidationException(string message, string fieldName, Exception innerException) : base(message, innerException)
        {
            FieldName = fieldName;
        }
    }
}
