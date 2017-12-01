using System;
using ClosedXML.Excel;
using System.Globalization;

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
            cellDateFull        = "J13",

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

            workbook.SaveAs(outFile);
        }
    }
}

