using System;
using System.IO;
using Aspose.Words;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Word;
using Break = DocumentFormat.OpenXml.Drawing.Break;
using Document = Microsoft.Office.Interop.Word.Document;
using PageSetup = Aspose.Words.PageSetup;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;

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
                DocumentFormat.OpenXml.Wordprocessing.Document doc = mainPart.Document;

                foreach (var paragraph in doc.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                {
                    var run = new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Break { Type = BreakValues.Page });
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
