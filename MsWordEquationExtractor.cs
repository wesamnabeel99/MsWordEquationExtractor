using System;

using EquationToImageConverter;

namespace Example
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputFile = @"C:\Users\Wesam Nabeel\x1.docx";
            string outputFile = "C:\\Users\\Wesam Nabeel\\equations.pdf";
            EquationConverter equationToImage = new EquationConverter(inputFile, outputFile);
            int extractedEquations = equationToImage.Convert();

            Console.WriteLine(extractedEquations + " equations extracted");
            Console.WriteLine("Press Any Key To Continue...");
            Console.ReadKey();

        }

    

    }


}
