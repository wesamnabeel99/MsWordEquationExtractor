using System;
using System.Drawing;
using EquationToImageConverter;
using Microsoft.Office.Interop.Word;
using WordEquationToImageConverter;

namespace PageSizeAdjustment
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            EquationSaver equationsSaver = new EquationSaver();
            equationsSaver.SelectAllEquations();

            EquationConverter equationConverter = new EquationConverter(@"C:\Users\Wesam Nabeel\equationsFile.docx", @"C:\Users\Wesam Nabeel\equationsFile.pdf", 10);
            equationConverter.Convert();

        }
    }
}
