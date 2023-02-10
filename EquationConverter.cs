using System;
using System.Drawing;
using System.Windows.Forms;
using Aspose.Pdf.Drawing;
using Aspose.Words;
using Aspose.Words.Math;
using iTextSharp.text.pdf;
using Microsoft.Office.Interop.Word;
using Document = iTextSharp.text.Document;
using Image = iTextSharp.text.Image;
using Paragraph = Aspose.Words.Paragraph;
using Rectangle = iTextSharp.text.Rectangle;
using Section = Aspose.Words.Section;

namespace EquationToImageConverter
{
    public class EquationConverter
    {
        private readonly string _inputFilePath;
        private readonly string _outputFilePath;
        private readonly int _quality;
        private int _equationNumber;

        public EquationConverter(string inputFilePath, string outputFilePath, int quality)
        {
            _inputFilePath = inputFilePath;
            _outputFilePath = outputFilePath;
            _quality = quality;
        }

        public int Convert()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(_inputFilePath);
            Document pdfDoc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new System.IO.FileStream(_outputFilePath, System.IO.FileMode.Create));
            pdfDoc.Open();

            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                NodeCollection equations = paragraph.GetChildNodes(NodeType.OfficeMath, false);
                foreach (OfficeMath equation in equations)
                {
                    Bitmap image = GetEquationImage(equation);

                    Console.WriteLine("equation " + (++_equationNumber) + ": " + equation.GetText());

                    Image iTextSharpImage = Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                    WriteToNewPdfPage(pdfDoc, iTextSharpImage);
                }
            }

            Console.WriteLine("done");
            Console.ReadKey(); 
            pdfDoc.Close();
            return _equationNumber;
        }



        private Bitmap GetEquationImage(OfficeMath equation)
        {
            SizeF size = GetImageSize(equation);

            int width = (int)Math.Ceiling(size.Width);
            int height = (int)Math.Ceiling(size.Height);

            Bitmap image = new Bitmap(width, height);

            Graphics graphics = Graphics.FromImage(image);

            equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, _quality);

            return image;
        }

        private Size GetImageSize(OfficeMath equation)
        {
            Bitmap placeHolderImage = new Bitmap(1, 1);
            Graphics placeHolderGraphics = Graphics.FromImage(placeHolderImage);
            SizeF size = equation.GetMathRenderer().RenderToScale(placeHolderGraphics, 0f, 0f, _quality);
            int width = (int)Math.Ceiling(size.Width);
            int height = (int)Math.Ceiling(size.Height);
            return new Size(width, height);
        }

        private void WriteToNewPdfPage(Document pdfDoc, Image iTextSharpImage)
        {
            pdfDoc.SetPageSize(new Rectangle(0, 0, iTextSharpImage.Width/_quality, iTextSharpImage.Height/_quality));
            pdfDoc.NewPage();
            iTextSharpImage.ScaleToFit(new Rectangle(0, 0, iTextSharpImage.Width/_quality, iTextSharpImage.Height/_quality));
            iTextSharpImage.SetAbsolutePosition(0, 0);
            pdfDoc.Add(iTextSharpImage);
        }
    }
}
