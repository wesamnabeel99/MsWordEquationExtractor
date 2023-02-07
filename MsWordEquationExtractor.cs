using System;
using WordEquationToImageConverter;

namespace PageSizeAdjustment
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            EquationSaver equationSaver = new EquationSaver();
            equationSaver.SaveEquation();
        }
    }
}
