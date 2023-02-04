using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Windows.Forms;
using System.Xml.XPath;
using Aspose.Words;
using Aspose.Words.Math;
using Microsoft.Office.Interop.Word;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(@"C:\Users\Wesam Nabeel\x1.docx");

            int equationNumber = 0;

            foreach (Aspose.Words.Paragraph pargraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                NodeCollection equations = pargraph.GetChildNodes(NodeType.OfficeMath, false);
                foreach (OfficeMath equation in equations)
                {
                    Bitmap placeHolderImage = new Bitmap(1, 1);
                    Graphics placeHolderGraphics = Graphics.FromImage(placeHolderImage);

                    SizeF size = equation.GetMathRenderer().RenderToScale(placeHolderGraphics, 0f, 0f, 2f);

                    int width = (int)Math.Ceiling(size.Width);
                    int height = (int)Math.Ceiling(size.Height);

                    Bitmap image = new Bitmap(width, height);

                    Graphics graphics = Graphics.FromImage(image);

                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 2f);

                    Console.WriteLine(equation.GetText() + " rendered with " + size.ToString());

                    Console.WriteLine("Horizontal Resolution is:" + image.HorizontalResolution.ToString() + "Vertical Resolution is:" + image.VerticalResolution.ToString());


                    image.Save("C:\\Users\\Wesam Nabeel\\equation" + equationNumber + ".png");
                    equationNumber++;
                }
            }

            Console.WriteLine("Press Any Key To Continue..");
            Console.ReadKey();
        }
    }
}