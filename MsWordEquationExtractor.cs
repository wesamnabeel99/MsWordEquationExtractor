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
            Aspose.Words.Document doc = new Aspose.Words.Document(@"C:\Users\Wesam Nabeel\x1.docx");

            int equationNumber = 0;

            Document pdfDoc = new Document();
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new System.IO.FileStream("C:\\Users\\Wesam Nabeel\\equations.pdf", System.IO.FileMode.Create));
            pdfDoc.Open();

            foreach (Aspose.Words.Paragraph pargraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                NodeCollection equations = pargraph.GetChildNodes(NodeType.OfficeMath, false);
                foreach (OfficeMath equation in equations)
                {
                    Bitmap image = new Bitmap(1, 1);
                    Graphics graphics = Graphics.FromImage(image);

                    SizeF size = equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 2f);

                    int width = (int)Math.Ceiling(size.Width);
                    int height = (int)Math.Ceiling(size.Height);

                    image = new Bitmap(width, height);
                    graphics = Graphics.FromImage(image);

                    equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 2f);

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
            Console.WriteLine(equationNumber + " equations extracted");
            Console.ReadKey();
        }
    }
}
