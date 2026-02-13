namespace InsercaoColaborador.Application.Interfaces
{
    public interface IProcessamentoService
    {
        string Nome { get; }
        void Executar();
    }
}
