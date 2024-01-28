using GemBox.Document;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
class Program
{
    static void Main()
    {
        Console.Write("Enter the PDF file path: ");
        string pdfFilePath = Console.ReadLine();

        Console.Write("Enter the Word file name (e.g., output.docx): ");
        string wordFileName = Console.ReadLine();

        // Pdf dan sug'urib olish
        string pdfText = ExtractTextFromPdf(pdfFilePath);

        // Write text to Word document
        CreateWordDocument(wordFileName, pdfText);
        Console.WriteLine($"Data extracted from PDF and written to Word file: {wordFileName}");
    }
    static string ExtractTextFromPdf(string pdfFilePath)
    {
        using (PdfReader pdfReader = new PdfReader(pdfFilePath))
        {
            using (PdfDocument pdfDocument = new PdfDocument(pdfReader))
            {
                StringWriter textWriter = new StringWriter();
                for (int pageNumber = 1; pageNumber <= pdfDocument.GetNumberOfPages(); pageNumber++)
                {
                    var strategy = new iText.Kernel.Pdf.Canvas.Parser.Listener.LocationTextExtractionStrategy();
                    PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
                    parser.ProcessPageContent(pdfDocument.GetPage(pageNumber));
                    textWriter.WriteLine(strategy.GetResultantText());
                }
                return textWriter.ToString();
            }
        }
    }
    static void CreateWordDocument(string wordFileName, string content)
    {
        ComponentInfo.SetLicense("FREE-LIMITED-KEY");
        var document = new GemBox.Document.DocumentModel();
        var section = new GemBox.Document.Section(document);
        document.Sections.Add(section);
        var paragraph = new GemBox.Document.Paragraph(document, content);
        section.Blocks.Add(paragraph);
        document.Save(wordFileName);
    }
}