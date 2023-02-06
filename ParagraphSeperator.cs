using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PageSetup = Aspose.Words.PageSetup;

namespace ParagraphSeparator
{
    public class ParagraphSeperatorClass 
    {
        private readonly string _inputFilePath;
        private readonly string _outputFilePath;

        public ParagraphSeperatorClass(string inputFilePath, string outputFilePath)
        {
            _inputFilePath = inputFilePath;
            _outputFilePath = outputFilePath;
        }

        public void Separate()
        {
            using (WordprocessingDocument document = WordprocessingDocument.Open(_inputFilePath, true))
            {
                MainDocumentPart mainPart = document.MainDocumentPart;
                Document doc = mainPart.Document;

                foreach (var paragraph in doc.Descendants<Paragraph>())
                {
                    var run = new Run(new Break { Type = BreakValues.Page });
                    paragraph.Append(run);
                }

                mainPart.Document.Save();
            }

            Aspose.Words.Document doc2 = new Aspose.Words.Document(_inputFilePath);
            PageSetup pageSetup = doc2.FirstSection.PageSetup;
            doc2.Save(_outputFilePath);
        }
    }
}
