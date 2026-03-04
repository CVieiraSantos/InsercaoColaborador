namespace InsercaoColaborador.Entities.Validation.Excecao
{
    public class TransacaoValidationException : Exception
    {
        public string FieldName { get; }
        public TransacaoValidationException(string message, string fildName) : base(message) 
        {
            FieldName = fildName;
        }

        public TransacaoValidationException(string message, string fieldName, Exception innerException) : base(message, innerException)
        {
            FieldName = fieldName;
        }
    }
}
