using System;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Words;
using Aspose.Words.Math;

namespace Example
{
    class MsWordEquationExtractor
    {
        static void Main(string[] args)
        {
            Document doc = new Document(@"C:\Users\Wesam Nabeel\x2.docx");

            int i = 0;
            foreach (OfficeMath math in doc.GetChildNodes(NodeType.OfficeMath, true))
            {
                Bitmap image = new Bitmap(500, 500);
                Graphics graphics = Graphics.FromImage(image);

                float v = math.GetMathRenderer().RenderToSize(graphics, 1.0f, 1.0f, 500, 500);

                image.Save("C:\\Users\\Wesam Nabeel\\equation" + i + ".png", ImageFormat.Png);
                i++;
            }
        }
    }
}
