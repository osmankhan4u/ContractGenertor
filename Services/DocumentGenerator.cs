using Xceed.Words.NET;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace ContractGeneratorBlazor.Services
{
    public interface IDocumentGenerator
    {
        Task SaveDocumentAsync(DocX doc, string filePath, string format);
    }

    public class DocumentGenerator : IDocumentGenerator
    {
        public async Task SaveDocumentAsync(DocX doc, string filePath, string format)
        {
            if (format.Equals("Word", StringComparison.OrdinalIgnoreCase))
            {
                doc.SaveAs(filePath);
                return;
            }

            // Extract plain text from DocX
            var text = doc.Text;

            // Generate PDF using QuestPDF
            var pdfBytes = GeneratePdfFromText(text);
            await File.WriteAllBytesAsync(filePath, pdfBytes);

        }

        private byte[] GeneratePdfFromText(string text)
        {
            return QuestPDF.Fluent.Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Margin(50);
                    page.Content()
                        .Text(text)
                        .FontSize(12);
                });
            }).GeneratePdf();
    }
    }
}
