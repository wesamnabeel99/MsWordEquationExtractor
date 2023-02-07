using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Math;
using Microsoft.Office.Interop.Word;
using Document = Microsoft.Office.Interop.Word.Document;

namespace WordEquationToImageConverter
{
    public class EquationSaver
    {
        private int _equationNumber;
       

        public EquationSaver()
        {

        }

        public void SaveEquation()
        {
            Application app = new Application();
            Document doc = app.Documents.Add();
            app.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow;
            app.Visible = true;

            string text = File.ReadAllText(@"C:\Users\Wesam Nabeel\filename.txt");

            Console.WriteLine(text);

            var range = doc.Range();
            range.Text = "The equation x^2 + y^2 is:";
            range.InsertParagraphAfter();


            var equationss = doc.OMaths.Add(range);
            equationss.Text = text;
            foreach(OMath omath in range.OMaths)
            {
                omath.BuildUp();
            }
            Console.WriteLine("modify the equation as needed");
            Console.ReadKey();

           
            doc.SaveAs2(@"C:\Users\Wesam Nabeel\testingIt.docx");
            doc.Close();
            app.Quit();
            // Load the word document
            Aspose.Words.Document asposeDoc = new Aspose.Words.Document(@"C:\Users\Wesam Nabeel\testingIt.docx");
            
            // Find all equations in the document
            foreach (Aspose.Words.Paragraph paragraph in asposeDoc.Document.GetChildNodes(NodeType.Paragraph, true))
            {
                NodeCollection equations = paragraph.GetChildNodes(NodeType.OfficeMath, false);
                foreach (OfficeMath equation in equations)
                {
                    Bitmap image = GetEquationImage(equation);

                    // Save the equation as an image
                    image.Save(@"C:\Users\Wesam Nabeel\equation" + _equationNumber + ".png", System.Drawing.Imaging.ImageFormat.Png);

                        string newText = "FLAG_DONE"; // TODO: remove this part
                        File.WriteAllText(@"C:\Users\Wesam Nabeel\filename.txt", newText);

                    }
                }
        }

        private Bitmap GetEquationImage(OfficeMath equation)
        {
            SizeF size = GetImageSize(equation);

            int width = (int)Math.Ceiling(size.Width);
            int height = (int)Math.Ceiling(size.Height);

            Bitmap image = new Bitmap(width, height);
    
            Graphics graphics = Graphics.FromImage(image);


            equation.GetMathRenderer().RenderToScale(graphics, 0f, 0f, 10);
            return image;
        }

        private SizeF GetImageSize(OfficeMath equation)
        {
            Bitmap placeHolderImage = new Bitmap(1, 1);
            Graphics placeHolderGraphics = Graphics.FromImage(placeHolderImage);
            SizeF size = equation.GetMathRenderer().RenderToScale(placeHolderGraphics, 0f, 0f, 10);
            return size;
        }
    
    }


}
