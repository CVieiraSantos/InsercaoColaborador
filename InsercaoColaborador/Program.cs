using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Application.Services;

var servicos = new List<IProcessamentoService>
{
    new ColaboradorProcessamentoService(),
    new TransacaoProcessamentoServiceRelacaoPagSemPendencias(),
    new TransacaoProcessamentoServicePendencias1()
};

Console.WriteLine("Selecione o tipo de processamento:\n");

for (int i = 0; i < servicos.Count; i++)
{
    Console.WriteLine($"{i + 1} - {servicos[i].Nome}");
}

Console.Write("\nOpção: ");
var input = Console.ReadLine();

if (!int.TryParse(input, out var escolha) ||
    escolha < 1 || escolha > servicos.Count)
{
    Console.WriteLine("Opção inválida.");
    return;
}

servicos[escolha - 1].Executar();
