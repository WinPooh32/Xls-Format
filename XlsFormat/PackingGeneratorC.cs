using System;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class PackingGeneratorC
    {
        XLWorkbook workbook;
        IXLWorksheet worksheet;

        public struct TemplateMap{
            public string cellDateShort;
            public string cellDateFull;
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
            cellDateShort       = "D8",
            cellDateFull        = "J13",

            cellCar             = "K3",
            cellVin             = "K4",
            cellDocs            = "K5",
            cellDriver          = "K6",
            cellDriverPassport  = "K7",

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
                worksheet = workbook.Worksheet(1);

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
            CarsTableC carsTbl
        ){
            var cellCar = worksheet.Cell(templateMap.cellCar);
            cellCar.SetValue (carsTbl.cars[0].name);

            workbook.SaveAs(outFile);
        }
    }
}

