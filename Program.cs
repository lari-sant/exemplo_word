using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Criação do documento
            //Cria um documento com o 
                Document exemploDoc = new Document();
            #endregion

            #region Criação de secção no Documento
            //Adiciona uma seção com o nome secaoCapa ao documento
            //Cada seção pode ser entendida como uma pagina do documento
                Section secaoCapa = exemploDoc.AddSection();
            #endregion

            #region Criar um paragrafo
            //Cria um paragrafo com o nome titulo e adiciona á seção secaoCapa
            //Os paragrafos são necessarios para insenção de textos, imagens, tabela etc
                Paragraph titulo = secaoCapa.AddParagraph();
            #endregion

            #region Adiciona texto ao paragrafo
            //Adiciona o texto Exemplo de Titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de titulo\n\n");
            #endregion 

            #region Formatar paragrafo 
            //Através da propriedade HorizontalAligment, é possivel alinhar o paragrafo
               titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
               ParagraphStyle estilo01 = new ParagraphStyle(exemploDoc);
               //Adiciona um nome ao estilo01
               estilo01.Name = "Cor do titulo";

               //Definir a cor do texto
               estilo01.CharacterFormat.TextColor = Color.DarkBlue;

               //Definir que o texto será em negrito
               estilo01.CharacterFormat.Bold = true;

               // Adiciona o estilo01 ao documento
               exemploDoc.Styles.Add(estilo01);

               //Aplica o estilo01 ao paragrafo titulo
               titulo.ApplyStyle(estilo01.Name);
            #endregion

            #region Trabalhar com tabulação
            //Adiciona um paragrafo 
            Paragraph textoCapa = secaoCapa.AddParagraph();

            textoCapa.AppendText("\tEste é um exemplo de texto com tabulação\n");

            Paragraph texto2Capa = secaoCapa.AddParagraph();

            texto2Capa.AppendText("\tBasicamento, então, uma seção representa uma pagina do documento e os paragrafos dentro de uma mesma seção" + "obviamente, aprecem na mesma pagina");
            #endregion

            #region Inserir imagem
               //Adiciona um paragrafo a seção secaoCapa
               Paragraph imagemCapa = secaoCapa.AddParagraph();
               
               //
               imagemCapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao documento\n\n");
            
               //Centraliza horizontalmente o paragrafo imagemCapa
               imagemCapa.Format.HorizontalAlignment = HorizontalAlignment.Center;
               
               //Adiciona imagem com o nome imagemExemplo ao paragrafo imagemCapa
               DocPicture imagemExemplo = imagemCapa.AppendPicture(Image.FromFile(@"saida\IMG\logo_csharp.png"));

               //Define uma largura e uma altura p a imagem

               imagemExemplo.Width = 300;
               imagemExemplo.Height = 300;

            #endregion

            #region Adicionar nova seção
            Section secaoCorpo = exemploDoc.AddSection();

            // Adiciona um paragrafo á seção secaoCorpo
            Paragraph paragraphCorpo = secaoCorpo.AddParagraph();
            
            //Adiciona um texto ao paragrafo paragraCorpo
            paragraphCorpo.AppendText("\tEste é um exemplo de paragrafo criado em uma nova seção." + "\tComo foi criada  uma nova seção, percebe que este texto aparece em uma nova pagina");
            
            #endregion

            #region  Adiciona uma tabela

            //Adiciona uma tabela a seção secaoCorpo
            Table tabela = secaoCorpo.AddTable(true);

            //Cria o cabeçalho da tabela
            string[] cabecalho = {"Item", "Descrição ", "Qtd. ", "Preço Unit. ","Preço"};

            string[][] dados = {
                new String[]{"Cenoura","Vegetal muito nutritivo","1", "R$4,00","R$ 4,00"},
                new String[]{"Batata","Vegetal muito consumido","2", "R$5,00","R$ 10,00"},
                new String[]{"Cenoura","Vegetal utilizado desde 500 a.C","1", "R$1,50","R$ 1,50"},
                new String[]{"Cenoura","Tomate é uma fruta","2", "R$6,00","R$ 12,00"},
            };

            //Adiciona as celulas
            tabela.ResetCells(dados.Length + 1, cabecalho.Length);
            
            //Adiciona uma linha na posição [0] do vetor de linhas
            //e define que esta linha é um cabeçalho
            TableRow Linha1 = tabela.Rows[0];
            Linha1.IsHeader = true;
            
            //Define a altura da linha 
            Linha1.Height = 23;

            //Formatação do cabeçalho
            Linha1.RowFormat.BackColor = Color.AliceBlue;

            // Percorre as linhas do cabeçalho
            for (int i = 0; i <cabecalho.Length; i++)
            {
                Paragraph p = Linha1.Cells[i].AddParagraph();
                Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                //Formatar dados do cabeçalho
                TextRange TR = p.AppendText(cabecalho[i]);
                TR.CharacterFormat.FontName = "Calibri";
                TR.CharacterFormat.FontSize = 14;
                TR.CharacterFormat.TextColor = Color.Teal;
                TR.CharacterFormat.Bold = true;
            } 

            //Adiciona as linhas do corpo da tabela 
            for (int r = 0; r < dados.Length; r++)
            {
                TableRow LinhaDados = tabela.Rows[r + 1];
                
                //Define a altura da Linha
                LinhaDados.Height = 28;

                //Percorre as colunas 
                for (int c = 0; c < dados[r].Length; c++)
                {
                   //Alinha as colunas
                   LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment.Middle;

                   //Preencher os dados nas linhas 
                   Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                   TextRange TR2 = p2.AppendText(dados[r][c]); 

                   //Formatar as celulas 
                   p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                   TR2.CharacterFormat.FontName = "Calibri";
                   TR2.CharacterFormat.FontSize = 12;
                   TR2.CharacterFormat.TextColor = Color.Brown;
                }
            }
            #endregion

            #region Salvar arquivo
            //Salva o arquivo em Docx
            //Utiliza o metodo SaveToFile p salvar o arquivo no formato desejado 
            //Assim como no word. caso já exista um arquivo com este nome, e substituido
                exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word_docx", FileFormat.Docx);
            #endregion


        }
    }
}
