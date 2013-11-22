using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddinTester
{
    class Program
    {
        static void Main(string[] args)
        {
            // Test the Smith-Wilson function
            var inputs = new double[,] {{1,0.95},{2,0.9},{3,0.85}};
            var alpha = 0.1;
            var ufr = 0.042;
            var tau = 20d;

            var calculatedBondPrice = ExcelAddin.ExcelMaths.SmithWilson(inputs, ufr, tau, alpha);
            Console.WriteLine("Bond price is: {0}", calculatedBondPrice);
            Console.ReadLine();
        }
    }
}
