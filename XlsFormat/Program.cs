using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Collections;

namespace XlsFormat
{
    class Common{
        private Common(){
        }

        static public IEnumerator<IXLCell> getCellsEnumerator(IXLWorksheet worksheet, string column){
            return worksheet.Column(column).CellsUsed().GetEnumerator();
        }
    }

    class MainClass
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var codesTable = new CodesTableC("/home/awake-monoblock/xlsx/Коды ТН ВЭД ОБЩАЯ база.xlsx");
            var batchTable = new BatchTableC("/home/awake-monoblock/xlsx/104 партия начальный формат.xlsx");
            var carsTable = new CarsTableC("/home/awake-monoblock/xlsx/ТранспортБД.xlsx");

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

            foreach (Car car in carsTable.cars) {
                Console.WriteLine (car.name + " " + car.docs + " " + car.vin);
            }

            //workbook.SaveAs("HelloWorld.xlsx");
        }
    }
}