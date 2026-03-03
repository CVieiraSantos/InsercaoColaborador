
namespace InsercaoColaborador.Entities.Validation.Excecao
{
    public class ContratoValidationException : ArgumentException
    {
        public string FieldName { get; }

        public ContratoValidationException(string message, string fieldName)
            : base(message)
        {
            FieldName = fieldName;
        }

        public ContratoValidationException(string message, string fieldName, Exception innerException)
            : base(message, innerException)
        {
            FieldName = fieldName;
        }

    }
}
