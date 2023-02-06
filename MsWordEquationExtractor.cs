using System;
using EquationToImageConverter;
using Microsoft.Office.Interop.Word;

namespace PageSizeAdjustment
{
    class Program
    {
        static void Main(string[] args)
        {
            DocxEditor docxEditor = new DocxEditor("C:\\Users\\Wesam Nabeel\\textdocx.docx");
            docxEditor.deleteParagraphs();
            docxEditor.findAndReplace("^p", "^m");
            Console.WriteLine("Done! press any key to save the file");
            Console.ReadKey();
            docxEditor.SaveDocument();
            docxEditor.CloseApp();
        }
    }
}
