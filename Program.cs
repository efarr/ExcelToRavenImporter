using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelToRavenImporter
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Count() == 2)
            {
                Import(args[0], args[1]);
            }
            else
            {
                Usage();
            }
        }

        private static void Import(string excelFileName, string ravenUrl)
        {
            new ExcelImporter(excelFileName, ravenUrl).Execute();
        }

        private static void Usage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("\tExcelToRavenImporter ExcelFilePath RavenDbURL");
        }
    }
}
