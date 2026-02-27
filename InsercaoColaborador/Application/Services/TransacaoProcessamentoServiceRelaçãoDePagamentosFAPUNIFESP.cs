using ClosedXML.Excel;
using InsercaoColaborador.Application.Interfaces;
using InsercaoColaborador.Entities.Transacao;
using InsercaoColaborador.Extension;
using InsercaoColaborador.Infrastructure.Sql.Builders;
using InsercaoColaborador.Infrastructure.Sql.Mappers.TransacaoMapper;
using InsercaoColaborador.Service;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace InsercaoColaborador.Application.Services
{
    public class TransacaoProcessamentoServiceRelaçãoDePagamentosFAPUNIFESP : IProcessamentoService
    {
        public string Nome => "Gerar INSERT de Transações FAPUNIFESP 01.2025 a 03.2025 - 04.25 a 09.25";

        public void Executar()
        {
            bool planilhaAbril = false;
            string caminhoExcel = @"C:\Users\Carlos Vieira\Downloads\Final 250826_Glosado_Relação_de_Pagamentos_FAPUNIFESP 01.2025 a 03.2025_NOVO.xlsx";
            string caminhoExcelAbril = @"C:\Users\Carlos Vieira\Downloads\Relação_de_Pagamentos_FAPUNIFESP 04.2025 a 09.2025.xlsx";
            string caminhoSql = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_novo_Relacao_de_pagamentos.sql";
            string caminhoSqlAbril = @"C:\Users\Carlos Vieira\Downloads\insert_transacoes_novo_Relacao_de_pagamentos_Abril.sql";

            List<string> caminhos = new List<string>() { caminhoExcel, caminhoExcelAbril };

            foreach(var caminhoPlanilha  in caminhos)
            {
                if (!File.Exists(caminhoPlanilha))
                {
                    Console.Error.WriteLine($"Arquivo Excel não encontrado!");
                    return;
                }

                var transacaoExcel = ExcelService.ImportarExcel(caminhoPlanilha, "Relação de pagamentos", linha =>
                {
                    return new TransacaoExcel
                    {
                        Item = linha.Cell(1).GetString().Trim(),
                        TipoDespesa = linha.Cell(3).GetString().Trim(),
                        SubCategoriaDeDespesa = linha.Cell(4).GetString().Trim(),
                        NomeCredor = linha.Cell(5).GetString().Trim(),
                        CnpjCpf = CpfCnpjGenerator.FormatarCpfOuCnpj(linha.Cell(6).GetString()),
                        ServicoProduto = linha.Cell(7).GetString().Trim(),
                        NumeroCheque = linha.Cell(8).GetString().Trim(),
                        DataPagamento = ParseExcelDate(linha.Cell(9)) ?? throw new FormatException($"Invalid data in row {linha.RowNumber()} column 7"),
                        Total = linha.Cell(10).GetDecimal(),
                        StatusAnalise = linha.Cell(11).GetString().Trim(),
                        Justificativa = linha.Cell(12).GetString().Trim(),
                        ValorGlosadoString = linha.Cell(14).GetString().Trim(),
                        ApuracaoGlosaParcial = linha.Cell(12).GetString().Trim(),
                    };
                });

                // quando for a planilha de abril a coluna justificativa será igual statusanalise que na planilha se refere a coluna Resultado
                // quando for a planilha de março 

                if (caminhoPlanilha.Contains("04.2025"))
                {
                    planilhaAbril = true;

                    foreach (var item in transacaoExcel)
                    {
                        item.Justificativa = item.StatusAnalise;
                    }
                }

                var transacoes = transacaoExcel.Select(e =>
                {
                    var DescricaoDespesa = string.Concat(e.TipoDespesa, " - ", e.SubCategoriaDeDespesa);
                    var nomeBeneficiario = GerarNomeBeneficiario.GetNomeBeneficiario(e);
                    int status = ResolverStatus(e.StatusAnalise, e);

                    return new Transacao
                    {

                        IdExtrato = 10113,
                        Numero = e.NumeroCheque.Length > 20 ? e.NumeroCheque.Substring(0, 20) : e.NumeroCheque,
                        DataTransacao = e.DataPagamento,
                        Descricao = DescricaoDespesa.Length > 200 ? DescricaoDespesa.Substring(0, 200) : DescricaoDespesa,
                        Valor = e.Total,
                        Tipo = "DEBIT",
                        NotaFiscal = string.IsNullOrWhiteSpace(e.ServicoProduto) ? string.Empty : e.ServicoProduto,
                        Categoria = "Outras despesas",
                        DataNotaFiscal = e.DataPagamento,
                        NomeBeneficiario = nomeBeneficiario.Length > 100 ? nomeBeneficiario.Substring(0, 100) : nomeBeneficiario,
                        OrigemRecurso = "Estadual",
                        IdParceria = 265,
                        Referencia = e.DataPagamento.Month,
                        Exercicio = e.DataPagamento.Year,
                        Status = status,
                        Avaliador = "Fernanda Biondi",
                        ValorContestado = ObterValorGlosaParcial(e, status), // e.ValorGlosado,
                        Conciliado = 1,
                        DataHoraConciliacao = DateTime.Now,
                        ObservacoesEntidade = planilhaAbril ? string.Concat("2.", e.Item) : e.Item,
                        ObservacoesOrgao = e.StatusAnalise, // e.Observacoes,
                        DataHoraCadastro = DateTime.Now,
                        IdCliente = 22,
                        IdEBanco = 164,
                        MeioPagamento = 1,
                        ValorDocumento = 0,
                        ValorEncargos = 0,
                        EstadoEmissor = 26,
                        SubCategoria = 0,
                        ExisteRateio = 0,
                        AnaliseEscrita = e.Justificativa.Length > 200 ? e.Justificativa.Substring(0,200) : e.Justificativa //e.StatusAnalise
                    };
                }).ToList();

                var sql = SqlInsertBuilder.BuildInsert(
                   table: "transacao",
                   columns: TransacaoSqlColumns.All,
                   items: transacoes,
                   valuesProjection: TransacaoSqlMapper.MapValues
               );

                if(!planilhaAbril)
                    File.WriteAllText(caminhoSql, sql, Encoding.UTF8);
                else
                    File.WriteAllText(caminhoSqlAbril, sql, Encoding.UTF8);

            }

            //Relação_de_Pagamentos_FAPUNIFESP 04.2025 a 09.2025


            DateTime? ParseExcelDate(IXLCell cell)
            {

                var s = cell.GetString()?.Trim();
                if (string.IsNullOrEmpty(s)) return null;

                if (DateTime.TryParse(s, new CultureInfo("pt-BR"), DateTimeStyles.None, out DateTime dta))
                    return dta;

                var formats = new[] { "M/d/yyyy", "M/d/yy", "MM/dd/yyyy", "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" };
                if (DateTime.TryParseExact(s, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                    return dt;
                if (DateTime.TryParse(s, new CultureInfo("en-US"), DateTimeStyles.None, out dt))
                    return dt;
                if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                    return dt;

                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    return DateTime.FromOADate(d);

                return null;
            }

            

            static int ResolverStatus(string? statusAnalise, TransacaoExcel transacaoExcel)
            {
                var status = statusAnalise?.Trim().ToLowerInvariant();
                var justificativa = transacaoExcel?.Justificativa?.Trim().ToLowerInvariant() ?? string.Empty;
                
                if (transacaoExcel == null || transacaoExcel.DataPagamento == default)
                    return 0;

                var mes = transacaoExcel.DataPagamento.Month;

                if (mes >= 1 && mes <=3)
                {
                    if (string.IsNullOrWhiteSpace(status))
                        return 1;
                    else if (status == "esclarecimento")
                        return 0;
                    else if (status == "esclarecido")
                        return 1;
                    else if (status == "glosado")
                        return 2;
                    else if (status == "devolvido")
                        return 0;

                    return 0;
                }
                else if (mes >= 4 && mes <= 9)
                {
                    if (status == "esclarecimento" || status == "esclarecimentos")
                        return 0;
                    if (string.IsNullOrWhiteSpace(status) || status == "ok")
                        return 1;
                    if (status == "glosada" || status == "glosado")
                        return 2;
                    if (status == "parcialmente glosada")
                        return 3;
                    
                    return 0;
                }

                return 0;
            }

            //static decimal? ObterValorGlosaParcial(TransacaoExcel transacaoExcel, int status)
            //{
            //    if(status != 3)
            //        return null;
            //    //tentar converter o valor glosado string pra decimal, se converter com sucesso, retorna ele.
            //    // Se não conseguir converter, vai ter que pegar a coluna L e isolar os números, e depois que isolar os números
            //    //converter pra decimal. Se converteu com sucesso, retorna o numero convertido, se nao, retorna nulo.
            //    var arrayDividido =  transacaoExcel.ApuracaoGlosaParcial.Split(" ");
            //    var valorString = arrayDividido.Last();
            //}

            static decimal? ObterValorGlosaParcial(TransacaoExcel transacaoExcel, int status)
            {
                // Validação de segurança
                if (status != 3)
                    return null;

                // Configuração para o padrão brasileiro (vírgula como decimal)
                var cultureBR = new CultureInfo("pt-BR");
                var estilo = NumberStyles.Number | NumberStyles.AllowCurrencySymbol;

                if (decimal.TryParse(transacaoExcel.ValorGlosadoString, estilo, cultureBR, out decimal valorGlosado))
                {
                    return valorGlosado;
                }

                // 1. TENTATIVA DIRETA: Pega a última palavra da frase
                var partes = transacaoExcel.ApuracaoGlosaParcial.Trim().Split(' ');
                var ultimaPalavra = partes.Last();

                if (decimal.TryParse(ultimaPalavra, estilo, cultureBR, out decimal valorDireto))
                {
                    return valorDireto;
                }

                // 2. ISOLAMENTO (Coluna L): Se falhar, limpa a string para extrair apenas o número
                // O padrão @"[^\d,.]" remove tudo que não for dígito, vírgula ou ponto
                string apenasNumeros = Regex.Replace(transacaoExcel.ApuracaoGlosaParcial, @"[^\d,.]", "");

                // Tenta converter a string limpa
                if (decimal.TryParse(apenasNumeros, estilo, cultureBR, out decimal valorIsolado))
                {
                    return valorIsolado;
                }

                // Retorna nulo se nenhuma das tentativas de conversão funcionar
                return null;
            }




        }
    }
}
