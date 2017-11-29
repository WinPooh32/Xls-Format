using System;
using ClosedXML.Excel;

namespace XlsFormat
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var codesTable = new CodesTableC("/home/awake-monoblock/Projects/XlsFormat/XlsFormat/bin/Debug/1.xlsx");

            for (int i = 0; i < codesTable.codes.Count; ++i){
                Console.WriteLine(codesTable.codes[i] + " " + codesTable.names[i]);
            }

            //workbook.SaveAs("HelloWorld.xlsx");
        }
    }
}