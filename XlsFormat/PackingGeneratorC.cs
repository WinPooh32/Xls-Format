using System;
using ClosedXML.Excel;
using System.Globalization;
using System.Data;
using System.Collections.Generic;

namespace XlsFormat
{
    public class PackingGeneratorC
    {
        XLWorkbook workbook;
        IXLWorksheet ws;

        public struct TemplateMap{
            public string cellDateShortFirst;
            public string cellDateShortSecond;
            public string cellDateFull;

            public string cellNumberFirst;
            public string cellNumberSecond;

            public string cellCar;
            public string cellVin;
            public string cellDocs;

            public string cellDriver;
            public string cellDriverPassport;

            public string columnId;
            public string columnName;
            public string columnMarkerCode;
            public string columnBagNumber;
            public string columnCode;

            public string columnAmount;
            public string columnUnitName;
            public string columnPlaces;

            public string columnPackageType;
            public string columnWeightGross;
            public string columnWeightNet;

            public string columnPrice;
            public string columnSumPrice;
        }

        private TemplateMap templateMap = new TemplateMap {
            cellDateShortFirst  = "A1",
            cellDateShortSecond = "D8",
            cellDateFull        = "J12",

            cellNumberFirst     = "A1",
            cellNumberSecond    = "D8", 

            cellCar             = "K3",
            cellVin             = "K4",
            cellDocs            = "K6",
            cellDriver          = "K7",
            cellDriverPassport  = "K8",

            columnId            = "A10",
            columnName          = "B10",
            columnMarkerCode    = "C10",
            columnBagNumber     = "D10",
            columnCode          = "E10",

            columnAmount        = "F10",
            columnUnitName      = "G10",
            columnPlaces        = "H10",

            columnPackageType   = "I10",
            columnWeightGross   = "J10",
            columnWeightNet     = "K10",

            columnPrice         = "L10",
            columnSumPrice      = "M10"
        };

        public PackingGeneratorC (string fileTemplate)
        {
            try{
                workbook = new XLWorkbook(fileTemplate);
                ws = workbook.Worksheet(1);

            }
            catch(Exception ex){
                Console.WriteLine(ex);
                throw new ArgumentException("[PackingGeneratorC] Error in file: " + fileTemplate);
            }
        }

        public void generatePackingList(
            string outFile, 
            BatchTableC batchTbl, 
            CodesTableC codesTbl, 
            Car car,
            Driver driver,
            string number
        ){
            string numberAndDate = number + " от " + DateTime.Now.ToString("dd.mm.yy.");

            string rawSecondShortDate     = Common.getCellString (ws, templateMap.cellDateShortSecond);
            string rawFullDate      = Common.getCellString (ws, templateMap.cellDateFull);
            string rawFirstNumber   = Common.getCellString (ws, templateMap.cellNumberFirst);

            string fullDate = DateTime.Now.ToString("dd MMMMM yyyy г.",  CultureInfo.CreateSpecificCulture("ru-RU"));

            Common.setCellString(ws, templateMap.cellDateFull,          rawFullDate.Replace("{fullDate}", fullDate));
            Common.setCellString(ws, templateMap.cellNumberFirst,       rawFirstNumber.Replace("{numberAndDate}", numberAndDate));


            string rawSecondNumber = Common.getCellString (ws,          templateMap.cellNumberFirst);
            Common.setCellString(ws, templateMap.cellNumberSecond,      rawSecondNumber.Replace("{numberAndDate}", numberAndDate));
          

            Common.setCellString(ws, templateMap.cellCar,               car.name);
            Common.setCellString(ws, templateMap.cellVin,               car.vin);
            Common.setCellString(ws, templateMap.cellDocs,              car.docs);

            Common.setCellString(ws, templateMap.cellDriver,            driver.name);
            Common.setCellString(ws, templateMap.cellDriverPassport,    driver.passport);


            //копируем футер в другое место и удаляем со старого места
            var footerCopy = ws.Range("A10:M16");

            ws.Cell (1, 16).Value = footerCopy;
            ws.Range("A10:M16").Delete(XLShiftDeletedCells.ShiftCellsUp);
            //------
            //вставляем таблицу со значениями
            var insertTable = GetTable (batchTbl, codesTbl);
            ws.Cell(9, 1).InsertTable(insertTable);

            //копируем футер на новое место с последующим удалением со старого
            footerCopy = ws.Range("P1:AB7");
            ws.Cell(insertTable.Rows.Count + 10, 1).Value = footerCopy;
            ws.Range("P1:AB7").Delete(XLShiftDeletedCells.ShiftCellsUp);


            ws.Cells ("A9:M9").Clear (XLClearOptions.Formats);

            workbook.SaveAs(outFile);
        }


        private DataTable GetTable(BatchTableC batchTbl, CodesTableC codesTbl){
            DataTable table = new DataTable();

            table.Columns.Add("№ П/П", typeof(UInt64));
            table.Columns.Add("Наименование", typeof(string));
            table.Columns.Add("Маркировка", typeof(string));
            table.Columns.Add("пломба", typeof(UInt64));
            table.Columns.Add("КОД ТНВЭД", typeof(UInt64));
            table.Columns.Add("Кол.", typeof(UInt32));
            table.Columns.Add("Ед.изм.", typeof(string));
            table.Columns.Add("Мест", typeof(Byte));
            table.Columns.Add("УПАКОВКА", typeof(string));
            table.Columns.Add("БРУТТО", typeof(double));
            table.Columns.Add("НЕТТО", typeof(double));
            table.Columns.Add("ЦЕНА", typeof(double));
            table.Columns.Add("СТОИМОСТЬ", typeof(double));

            var codes = codesTbl.codes;

            UInt32 i = 0;

            foreach (KeyValuePair<string, List<XlsFormat.BatchTableC.Product>> entry in batchTbl.goods) {
                var key = entry.Key;

                foreach (XlsFormat.BatchTableC.Product value in entry.Value) {
                    ++i;

                    ulong code;

                    if (!codes.TryGetValue (value.name, out code)) {
                        code = 0;
                    }

                    var netWeight = (value.bagWeight / value.allPlaces) * value.placesByType;

                    //FIXME надр нормально распределить
                    //             № П/П   Наименование    Маркировка  пломба           КОД ТНВЭД   Кол.             Ед.изм. Мест  УПАКОВКА    БРУТТО           НЕТТО   ЦЕНА         СТОИМОСТЬ
                    table.Rows.Add(i,      value.name,     key,        value.bagNumber, code,       value.placesByType, "шт.",  0,    "мешок",    value.bagWeight, netWeight,      value.price, value.price * value.placesByType);
                }
            }

            return table;
        }
    }
}