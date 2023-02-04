using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Math;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Document = iTextSharp.text.Document;
using Image = iTextSharp.text.Image;
using Rectangle = iTextSharp.text.Rectangle;

namespace EquationToImageConverter
{
    public class EquationConverter
    {
        private readonly string _inputFilePath;
        private readonly string _outputFilePath;
        private int _equationNumber;

        public EquationConverter(string inputFilePath, string outputFilePath)
        {
            _inputFilePath = inputFilePath;
            _outputFilePath = outputFilePath;
        }

        public int Convert()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(_inputFilePath);
            Document pdfDoc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new System.IO.FileStream(_outputFilePath, System.IO.FileMode.Create));
            pdfDoc.Open();

            foreach (Aspose.Words.Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                NodeCollection equations = paragraph.GetChildNodes(NodeType.OfficeMath, false);
                foreach (OfficeMath equation in equations)
                {
                    Bitmap image = GetEquationImage(equation);

                    Console.WriteLine("equation " + (++_equationNumber) + ": " + equation.GetText());

                    Image iTextSharpImage = Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                    WriteToNewPdfPage(pdfDoc,iTextSharpImage);
                }
            }

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

            equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 2f);

            return image;
        }

        private Size GetImageSize(OfficeMath equation)
        {
            Bitmap placeHolderImage = new Bitmap(1, 1);
            Graphics placeHolderGraphics = Graphics.FromImage(placeHolderImage);
            SizeF size = equation.GetMathRenderer().RenderToScale(placeHolderGraphics, 0f, 0f, 2f);
            int width = (int)Math.Ceiling(size.Width);
            int height = (int)Math.Ceiling(size.Height);
            return new Size(width, height);
        }

        private void WriteToNewPdfPage(Document pdfDoc, Image iTextSharpImage)
        {
            pdfDoc.SetPageSize(new Rectangle(0, 0, iTextSharpImage.Width, iTextSharpImage.Height));
            pdfDoc.NewPage();
            iTextSharpImage.SetAbsolutePosition(0, 0);
            pdfDoc.Add(iTextSharpImage);
        }
    }
}
