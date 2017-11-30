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
            var batchTable = new BatchTableC("/home/awake-monoblock/xlsx/104 партия начальный формат.xlsx");

//            foreach(KeyValuePair<string, UInt64> entry in codesTable.codes)
//            {
//                Console.WriteLine (entry.Key + " " + entry.Value);
//                // do something with entry.Value or entry.Key
//            }

            foreach(KeyValuePair<string, List<XlsFormat.BatchTableC.Product>> entry in batchTable.goods)
            {
                foreach (XlsFormat.BatchTableC.Product value in entry.Value) {
                    Console.WriteLine (entry.Key + " " + value.price);
                }
            }

            codesTable = null;
            //workbook.SaveAs("HelloWorld.xlsx");
        }
    }
}