using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Text;

Console.WriteLine("Extração de textos de PDF");

//string dirBase = AppDomain.CurrentDomain.BaseDirectory;
string dirBase = @"C:\Desenv\XMLs IA\Base classificada\";

Console.WriteLine(string.Format("Obtendo informações em diretório {0}", dirBase));

// Verifica se o arquivo Excel já existe, se não existir, cria um novo
string arqExcel = Path.Combine(dirBase, string.Format("DataNFSe_{0}.xls", DateTime.Now.ToString("yyyyMMddHHmmss")));

// Criar uma nova planilha XLS
IWorkbook workbook = new HSSFWorkbook();
ISheet sheet = workbook.CreateSheet("Dados");

// Cabeçalho da planilha
IRow headerRow = sheet.CreateRow(0);
headerRow.CreateCell(0).SetCellValue("Nome do Arquivo");
headerRow.CreateCell(1).SetCellValue("Conteúdo Extraído");

int linhaAtual = 1;
foreach (var arquivo in Directory.EnumerateFiles(dirBase, "*.pdf"))
{
    try
    {
        //Abrir o arquivo PDF
        string txtPdf = string.Empty;        
        using (PdfReader reader = new PdfReader(arquivo))
        {
            //Obter o documento PDF
            PdfDocument pdf = new PdfDocument(reader);

            //Extrair texto de cada página        
            for (int pagina = 1; pagina <= pdf.GetNumberOfPages(); pagina++)
            {
                ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                try
                {
                    txtPdf = txtPdf + PdfTextExtractor.GetTextFromPage(pdf.GetPage(pagina), strategy);
                }
                catch
                {
                    txtPdf = "";
                    pdf.Close();
                    reader.Close();
                    break;
                }
            }

            //Verificando conteúdo
            if (string.IsNullOrEmpty(txtPdf))
                continue;

            //Finalizando arquivo
            pdf.Close();
            reader.Close();
        }

        // Adicionar informações à planilha
        linhaAtual++;
        IRow dataRow = sheet.CreateRow(linhaAtual);
        dataRow.CreateCell(0).SetCellValue(Path.GetFileName(arquivo));
        dataRow.CreateCell(1).SetCellValue(txtPdf);
    }
    catch (Exception ex) { }
}

// Auto redimensionar as colunas para ajustar o conteúdo
for (int i = 0; i < 2; i++)
{
    sheet.AutoSizeColumn(i);
}

// Salvar o arquivo XLS com encoding UTF-8
using (FileStream file = new FileStream(arqExcel, FileMode.Create, FileAccess.Write))
{
    workbook.Write(file);
}

Console.WriteLine($"Arquivo Excel '{arqExcel}' gerado com sucesso. Pressione qualquer tecla para sair.");
Console.ReadKey();

