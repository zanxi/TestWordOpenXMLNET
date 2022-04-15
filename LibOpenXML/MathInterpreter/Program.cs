using Latex2MathML;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MathInterpreter
{
    class Program
    {
        static void Main(string[] args)
        {

            // {\pi_{ij}}={\rm{R}}SBU_{ij}^L-{\rm{R}}SBU_{ij}^H={q_{ap}}(f_{ij}^L-f_{ij}^H)/{N_{ap}}
            Console.WriteLine("\n\n\n\n\n\n\n --------------- Begin Create MathLm ---------------");

            string sourcePath = @"d:\_2022___\Src_3__\MathInterpreter\MathInterpreter\sourcePath.txt";
            string outputPath = @"d:\_2022___\Src_3__\MathInterpreter\MathInterpreter\outputPath.htm";

            //var lmm = new LatexToMathMLConverter(sourcePath, Encoding.GetEncoding(1251), outputPath);
            //lmm.ValidateResult = true;
            //lmm.ValidateResult = false;
            //lmm.Convert();

            //System.Diagnostics.Process.Start(outputPath);

            //Console.WriteLine("Create MathLm. Press key!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
            //System.Threading.Thread.Sleep(1000);
            //Console.ReadKey();

        }
    }
}
