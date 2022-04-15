using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;


using LibOpenXml;
using LibOpenXML;

namespace OpenXML_HeaderFooter
{
    class Program
    {
        static void Main(string[] args)
        {
            //TestLatexToWord lat = new TestLatexToWord();
            //lat.Main();

            TestGenerateDocument td = new TestGenerateDocument();
            td.GenerateWordDoc();

            //TestGenerateDocument td = new TestGenerateDocument();
            //td.GenerateWordDocMerge();
        }
    }
}
