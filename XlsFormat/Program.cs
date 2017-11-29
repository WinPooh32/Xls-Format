using System;
using ClosedXML.Excel;
using System.Collections.Generic;

namespace XlsFormat
{
    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var codesTable = new CodesTableC("/home/awake-monoblock/xlsx/Коды ТН ВЭД ОБЩАЯ база.xlsx");
           

            foreach(KeyValuePair<string, UInt64> entry in codesTable.codes)
            {
                Console.WriteLine (entry.Key + " " + entry.Value);
                // do something with entry.Value or entry.Key
            }
               
            //workbook.SaveAs("HelloWorld.xlsx");
        }
    }
}