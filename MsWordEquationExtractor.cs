using System;
using Microsoft.Office.Interop.Word;

namespace PageSizeAdjustment
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            Document doc = app.Documents.Open("C:\\Users\\Wesam Nabeel\\good test.docx");
            int paraCount = doc.Paragraphs.Count;

            for (int i = paraCount; i >= 1; i--)
            {
                Paragraph para = doc.Paragraphs[i];
                OMaths equations = para.Range.OMaths;
                if (equations.Count == 0)
                {

                    try
                    {
                        para.Range.Delete();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("couldn't delete " + para.Range.Text + " due to " + e);
                    }
                }
                
            }
            Console.ReadKey();
            doc.Save();
            app.Quit();


        }
    }
}
