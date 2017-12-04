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

        static public string getCellString (IXLWorksheet ws, string cell){
            return ws.Cell (cell).GetString();
        }

        static public void setCellString (IXLWorksheet ws, string cell, string value){
            ws.Cell (cell).SetValue (value);
        }
    }

    class MainClass
    {
        public static void Main(string[] args)
        {
            var codesTable = new CodesTableC("/home/awake-monoblock/xlsx/Коды ТН ВЭД ОБЩАЯ база.xlsx");
            var batchTable = new BatchTableC("/home/awake-monoblock/xlsx/104 партия начальный формат.xlsx");
            var carsTable = new CarsTableC("/home/awake-monoblock/xlsx/ТранспортБД.xlsx");
            var generatorPacking = new PackingGeneratorC("/home/awake-monoblock/xlsx/шаблоны/Упаковочный лист.xlsx");

            generatorPacking.generatePackingList(
                "/home/awake-monoblock/out.xlsx", 
                batchTable, codesTable, 
                carsTable.cars[0], 
                carsTable.drivers[0],
                "NOMER@12738"
            );

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