using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Math;
using iTextSharp.text.pdf;
using Document = iTextSharp.text.Document;
using Image = iTextSharp.text.Image;
using Rectangle = iTextSharp.text.Rectangle;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = @"C:\Users\Wesam Nabeel\x1.docx";
            string outputFile = "C:\\Users\\Wesam Nabeel\\equations.pdf";
            int extractedEquations = ExtractEquationsToPDF(inputFile, outputFile);

            Console.WriteLine(extractedEquations + " equations extracted");
            Console.WriteLine("Press Any Key To Continue...");
            Console.ReadKey();

        }

        private static int ExtractEquationsToPDF(string inputFile, string outputFile)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(inputFile);

            int equationNumber = 0;

            Document pdfDoc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new System.IO.FileStream(outputFile, System.IO.FileMode.Create));
            pdfDoc.Open();

            foreach (Aspose.Words.Paragraph pargraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                foreach (OfficeMath equation in pargraph.GetChildNodes(NodeType.OfficeMath, false))
                {
                    Bitmap image = GetEquationImage(equation);
                    Console.WriteLine("equation " + (equationNumber + 1) + ": " + equation.GetText());

                    Image iTextSharpImage = Image.GetInstance(image, System.Drawing.Imaging.ImageFormat.Png);
                    pdfDoc.SetPageSize(new Rectangle(0, 0, iTextSharpImage.Width, iTextSharpImage.Height));
                    pdfDoc.NewPage();
                    iTextSharpImage.SetAbsolutePosition(0, 0);
                    pdfDoc.Add(iTextSharpImage);

                    equationNumber++;
                }
            }

            pdfDoc.Close();
            return equationNumber;
        }

        private static Bitmap GetEquationImage(OfficeMath equation)
        {
            SizeF size = GetImageSize(equation);

            int width = (int)Math.Ceiling(size.Width);
            int height = (int)Math.Ceiling(size.Height);

            Bitmap image = new Bitmap(width, height);

            Graphics graphics = Graphics.FromImage(image);

            equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 2f);

            return image;
        }

        private static SizeF GetImageSize(OfficeMath equation)
        {
            Bitmap placeHolderImage = new Bitmap(1, 1);
            Graphics placeHolderGraphics = Graphics.FromImage(placeHolderImage);

            return equation.GetMathRenderer().RenderToScale(placeHolderGraphics, 0f, 0f, 2f);

        }
    }


}
