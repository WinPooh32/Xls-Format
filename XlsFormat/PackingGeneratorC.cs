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

        public int generatePackingList(
			string savePath, 
            BatchTableC batchTbl, 
            CodesTableC codesTbl, 
            Car car,
            Driver driver,
            string number
        ){
            string numberAndDate = number + " от " + DateTime.Now.ToString("dd.MM.yy.");

            string rawSecondShortDate     = Common.getCellString (ws, templateMap.cellDateShortSecond);
            string rawFullDate      = Common.getCellString (ws, templateMap.cellDateFull);
            string rawFirstNumber   = Common.getCellString (ws, templateMap.cellNumberFirst);

            string fullDate = DateTime.Now.ToString("dd MMMM yyyy г.",  CultureInfo.CreateSpecificCulture("ru-RU"));

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
			const string FOOTER_RANGE = "A10:M16";
            var footerCopy = ws.Range(FOOTER_RANGE);

            ws.Cell (1, 16).Value = footerCopy; //в незанятое место
			footerCopy.Clear(XLClearOptions.ContentsAndFormats);
            //------
            //вставляем таблицу со значениями
			var insertTable = GetTable (batchTbl, codesTbl, savePath);

			if (insertTable != null)
			{
				ws.Cell(9, 1).InsertTable(insertTable);
			}
			else
			{
				return 1;
			}

			int lastItemPos = insertTable.Rows.Count + 9;

			//копируем футер на новое место с последующим удалением со старого
			const string TMP_RANGE = "P1:AB7";
            footerCopy = ws.Range(TMP_RANGE);
			ws.Cell(lastItemPos + 1, 1).Value = footerCopy;
            ws.Range(TMP_RANGE).Clear(XLClearOptions.ContentsAndFormats);

			//ws.Cells ("A9:M9").Clear (XLClearOptions.Formats);

			ws.Range("J10:K" + lastItemPos + 1).Style.NumberFormat.Format = "0.000";

			Common.Log("Сохранение файлов в '" + savePath + "'");
			workbook.SaveAs(savePath + "\\" + Common.fileParty);

			return 0;
        }


        private UInt32 CalcItemsCount(BatchTableC batchTbl){
            UInt32 count = 0;
			var enumer = batchTbl.sortedProducts.GetEnumerator();
			while (enumer.MoveNext()) ++count;
            return count;
        }

        private DataTable GetTable(BatchTableC batchTbl, CodesTableC codesTbl, string savePath){
            DataTable table = new DataTable();

            table.Columns.Add("№ П/П", typeof(UInt64));
            table.Columns.Add("Наименование", typeof(string));
            table.Columns.Add("Маркировка", typeof(string));
            table.Columns.Add("Пломба", typeof(UInt64));
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
			var notfoundNames = new HashSet<string>();

			UInt32 itemsCount = CalcItemsCount(batchTbl);
			double packageWeight = (double)batchTbl.weightPackage / (double)itemsCount;

            UInt32 i = 0;

			ulong prevBagNumber = 0;

			double testTotalGross = 0;
			double testTotalNet = 0;

			foreach (XlsFormat.BatchTableC.Product value in batchTbl.sortedProducts)
			{
				++i;

				int place = (prevBagNumber != value.bagNumber ? 1 : 0);
				prevBagNumber = value.bagNumber;

                ulong code;
                if (!codes.TryGetValue (value.name, out code)) {
                    code = 0;
					notfoundNames.Add(value.name);
                }

				var netWeight = (value.bagWeight / (double)value.allPlaces) * (double)value.placesByType;

				var grossWeight = netWeight + packageWeight; //* (double)value.placesByType

				testTotalNet += netWeight;
				testTotalGross += grossWeight;

                //             № П/П   Наименование    Маркировка  пломба           КОД ТНВЭД   Кол.             Ед.изм. Мест  УПАКОВКА    БРУТТО           НЕТТО   ЦЕНА         СТОИМОСТЬ
				table.Rows.Add(i,      value.name,     value.number,        value.bagNumber, code,       value.placesByType, "шт.",  place,    "мешок",    grossWeight, netWeight,      value.price, value.price * value.placesByType);
            }

			if (notfoundNames.Count != 0)
			{
				Common.Log("Добавляем ненайденные коды ТН ВЭД");
				codesTbl.AppendNotFoundNames(notfoundNames);
				return null;
			}

			Common.Log(testTotalNet + " " + testTotalGross + " " + (testTotalGross - testTotalNet));

            return table;
        }
    }
}