using System;
using ClosedXML.Excel;
using System.Globalization;
using System.Data;
using System.Collections.Generic;
using System.Linq;

namespace XlsFormat
{
	public class PackingGeneratorC
	{

		//№ П/П Наименование    Маркировка пломба           КОД ТНВЭД  Кол.Ед.изм.Мест УПАКОВКА    БРУТТО НЕТТО      ЦЕНА СТОИМОСТЬ
		public struct BidloProduct
		{
			public UInt32 uid;

			public ulong code;
			public string name;
			public UInt64 bagNumber;

			public UInt32 places;
			public double grossWeight;
			public double netWeight;

			public decimal price;
		}

		List<BidloProduct> bidloProducts = new List<BidloProduct>();

		public struct TemplateMap
		{
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
		}

		public struct CMR_TemplateMap
		{
			public string carNumberAndDate_0;

			public string senderCity_4;
			public string dayAndMonth_4;
			public string year_4;

			public string specificationNumber_5;
			public string shortDate_spec_5;

			public string totalPlaces_6;
			public string totalGrossWeight_11;
			public string totalPrice_13;

			public string driverName_23;
			public string driverPassport_23;

			public string carNumberFull_25;
			public string carBrand_26;
		}

		public CMR_TemplateMap cmrTemplateMap = new CMR_TemplateMap
		{
			carNumberAndDate_0 = "AF3",

			senderCity_4 = "D19",
			dayAndMonth_4 = "D21",
			year_4 = "J21",

			specificationNumber_5 = "J23",
			shortDate_spec_5 = "P23",

			totalPlaces_6 = "B26",
			totalGrossWeight_11 = "AB26",
			totalPrice_13 = "J35",

			driverName_23 = "R44",
			driverPassport_23 = "O46",

			carNumberFull_25 = "B49",
			carBrand_26 = "K49"
		};

		private TemplateMap templatePackageMap = new TemplateMap {
            cellDateShortFirst  = "A1",
            cellDateShortSecond = "D8",
            cellDateFull        = "J12",

            cellNumberFirst     = "A1",
            cellNumberSecond    = "D8", 

            cellCar             = "K3",
            cellVin             = "K4",
            cellDocs            = "K6",
            cellDriver          = "K7",
            cellDriverPassport  = "K8"
        };

		private TemplateMap templateSpecificationMap = new TemplateMap {
            cellDateShortFirst  = "A1",
            cellDateShortSecond = "D8",
            cellDateFull        = "J14",

            cellNumberFirst     = "A1",
            cellNumberSecond    = "D8", 

            cellCar             = "J3",
            cellVin             = "J4",
            cellDocs            = "J6",
            cellDriver          = "J7",
            cellDriverPassport  = "J8"
        };

		double totalGross;
		double totalNet;
		UInt32 totalPlaces;
		decimal totalPrice;

		private void replaceTemplateValues(IXLWorksheet ws, 
		                              TemplateMap templateMap, 
		                              Car car,
		                              Driver driver,
		                              string number)
		{
			string numberAndDate = number + " от " + DateTime.Now.ToString("dd.MM.yy.");

			string rawSecondShortDate = Common.getCellString(ws, templateMap.cellDateShortSecond);
			string rawFullDate = Common.getCellString(ws, templateMap.cellDateFull);
			string rawFirstNumber = Common.getCellString(ws, templateMap.cellNumberFirst);

			string fullDate = DateTime.Now.ToString("dd MMMM yyyy г.", CultureInfo.CreateSpecificCulture("ru-RU"));

			Common.setCellString(ws, templateMap.cellDateFull, rawFullDate.Replace("{fullDate}", fullDate));
			Common.setCellString(ws, templateMap.cellNumberFirst, rawFirstNumber.Replace("{numberAndDate}", numberAndDate));


			string rawSecondNumber = Common.getCellString(ws, templateMap.cellNumberFirst);
			Common.setCellString(ws, templateMap.cellNumberSecond, rawSecondNumber.Replace("{numberAndDate}", numberAndDate));


			Common.setCellString(ws, templateMap.cellCar, car.name + ": " + car.number);
			Common.setCellString(ws, templateMap.cellVin, car.vin);
			Common.setCellString(ws, templateMap.cellDocs, car.docs);

			Common.setCellString(ws, templateMap.cellDriver, driver.name);
			Common.setCellString(ws, templateMap.cellDriverPassport, driver.passport);
		}

		private void moveFooterTo(IXLWorksheet ws, string rangeFrom, int rowTo, int columnTo)
		{	
			var footerCopy = ws.Range(rangeFrom);
			ws.Cell(rowTo, columnTo).Value = footerCopy;
			footerCopy.Clear(XLClearOptions.ContentsAndFormats);
		}

        public int generatePackingList(
			string template,
			string savePath, 
            BatchTableC batchTbl, 
            CodesTableC codesTbl, 
            Car car,
            Driver driver,
            string number
        ){

			XLWorkbook workbook;
			IXLWorksheet ws;

			try{
                workbook = new XLWorkbook(template);
				ws = workbook.Worksheet(1);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				throw new ArgumentException("[PackingGeneratorC] Error in file: " + template);
			}

			replaceTemplateValues(ws, templatePackageMap, car, driver, number);

			//копируем футер в другое место и удаляем со старого места
			moveFooterTo(ws, "A10:M16", 1, 16);

            //------
            //вставляем таблицу со значениями
			var insertTable = GetTable (batchTbl, codesTbl, savePath);

			if (insertTable != null)
			{
				ws.Cell(9, 1).InsertTable(insertTable);
			}
			else return 1;

			int lastItemPos = insertTable.Rows.Count + 9;

			//копируем футер на новое место с последующим удалением со старого
			moveFooterTo(ws, "P1:AB7", lastItemPos + 1, 1);

			ws.Range("J10:K" + lastItemPos + 1).Style.NumberFormat.Format = "0.000";

			Common.Log("Сохранение Упаковочник в '" + savePath + "'");
			workbook.SaveAs(savePath + "\\" + Common.fileParty);

			return 0;
        }

		public int GenerateSpecification(
			string template,
			string savePath, 
			BatchTableC batchTbl, 
			CodesTableC codesTbl,
			HashSet<UInt64> excludeList,
			Car car,
			Driver driver,
			string number
		)
		{
			XLWorkbook workbook;
			IXLWorksheet ws;

			try{
                workbook = new XLWorkbook(template);
				ws = workbook.Worksheet(1);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				throw new ArgumentException("[PackingGeneratorC] Error in file: " + template);
			}

			replaceTemplateValues(ws, templateSpecificationMap, car, driver, number);

			//копируем футер в другое место и удаляем со старого места
			moveFooterTo(ws, "A10:K21", 1, 16);

			//------
			//вставляем таблицу со значениями
			var insertTable = GetSpecificationTable(batchTbl, codesTbl, excludeList);
			ws.Cell(9, 1).InsertTable(insertTable);

			int lastItemPos = insertTable.Rows.Count + 9;

			//копируем футер на новое место с последующим удалением со старого
			moveFooterTo(ws, "P1:AA12", lastItemPos + 1, 1);

			ws.Range("I10:J" + lastItemPos + 1).Style.NumberFormat.Format = "0.000";
			ws.Range("K10:K" + lastItemPos + 1).Style.NumberFormat.Format = "0.00";

			Common.Log("Сохранение Спецификация в '" + savePath + "'");
			workbook.SaveAs(savePath + "\\" + Common.fileSpecification);

			return 0;
		}

		public void GenerateCMR(
			string template,
			string savePath,
			BatchTableC batchTbl,
			CodesTableC codesTbl,

			Car car,
			Driver driver,
			string number,
		
			string city)
		{
			XLWorkbook workbook;
			IXLWorksheet ws;

			try
			{
				workbook = new XLWorkbook(template);
				ws = workbook.Worksheet(1);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				throw new ArgumentException("[PackingGeneratorC] Error in file: " + template);
			}

			Common.setCellString(ws, cmrTemplateMap.carNumberAndDate_0, car.numberShort + DateTime.Now.ToString("ddMM"));

			Common.setCellString(ws, cmrTemplateMap.senderCity_4, city);
			Common.setCellString(ws, cmrTemplateMap.dayAndMonth_4, DateTime.Now.ToString("dd MMMM", CultureInfo.CreateSpecificCulture("ru-RU")));
			Common.setCellString(ws, cmrTemplateMap.year_4, DateTime.Now.ToString("yyyyг.", CultureInfo.CreateSpecificCulture("ru-RU")));

			Common.setCellString(ws, cmrTemplateMap.specificationNumber_5, number);
			Common.setCellString(ws, cmrTemplateMap.shortDate_spec_5, DateTime.Now.ToString("dd.MM.yy."));

			Common.setCellString(ws, cmrTemplateMap.totalPlaces_6, totalPlaces + " мешка п/п (интернет заказы для личного пользования)");
			Common.setCellString(ws, cmrTemplateMap.totalGrossWeight_11, totalGross.ToString("0.000"));
			Common.setCellString(ws, cmrTemplateMap.totalPrice_13, totalPrice.ToString());

			Common.setCellString(ws, cmrTemplateMap.driverName_23, driver.name);
			Common.setCellString(ws, cmrTemplateMap.driverPassport_23, "пас. " + driver.passport);

			Common.setCellString(ws, cmrTemplateMap.carNumberFull_25, car.number);
			Common.setCellString(ws, cmrTemplateMap.carBrand_26, car.name);

			Common.Log("Сохранение СМР в '" + savePath + "'");
			workbook.SaveAs(savePath + "\\" + Common.fileCMR);
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

			totalGross = 0;
			totalNet = 0;
			totalPrice = new decimal(0.0);

			foreach (XlsFormat.BatchTableC.Product value in batchTbl.sortedProducts)
			{
				++i;

				int place = (prevBagNumber != value.bagNumber ? 1 : 0);
				prevBagNumber = value.bagNumber;

				if (place == 1)
				{
					++totalPlaces;
				}

                ulong code;
                if (!codes.TryGetValue (value.name, out code)) {
                    code = 0;
					notfoundNames.Add(value.name);
                }

				var netWeight = (value.bagWeight / (double)value.allPlaces) * (double)value.placesByType;
				var grossWeight = netWeight + packageWeight; //* (double)value.placesByType
				var sumPrice = value.price * value.placesByType;

				totalNet += netWeight;
				totalGross += grossWeight;
				totalPrice += sumPrice;

                //             № П/П   Наименование    Маркировка    пломба           КОД ТНВЭД  Кол.                Ед.изм.  Мест   УПАКОВКА    БРУТТО       НЕТТО      ЦЕНА         СТОИМОСТЬ
				table.Rows.Add(i,      value.name,     value.number, value.bagNumber,      code, value.placesByType,  "шт.",  place,  "мешок",   grossWeight, netWeight, value.price, sumPrice);

				bidloProducts.Add(new BidloProduct
				{
					uid = i,

					bagNumber = value.bagNumber,
					name = value.name,
					code = code,

					places = value.placesByType,
					netWeight = netWeight,
					grossWeight = grossWeight,

					price = sumPrice
				});
			}

			Common.Log("Уникальных мешков в упаковочнике: " + totalPlaces);

			if (notfoundNames.Count != 0)
			{
				Common.Log("Добавляем ненайденные коды ТН ВЭД");
				codesTbl.AppendNotFoundNames(notfoundNames);
				return null;
			}

			Common.Log(totalNet + " " + totalGross + " " + (totalGross - totalNet));

            return table;
        }

		public BidloProduct AppendList(BidloProduct previtem, BidloProduct current, List<BidloProduct> fixedlist)
		{
			if (previtem.bagNumber == current.bagNumber &&
				previtem.code == current.code
		   	)
			{
				previtem.grossWeight +=current.grossWeight;
				previtem.price += current.price;
				previtem.places += current.places;
				previtem.netWeight += current.netWeight;
			}
			else {
				previtem = current;
			}

			fixedlist.Add(previtem);

			return previtem;
		}

		public DataTable GetSpecificationTable(BatchTableC batchTbl, CodesTableC codesTbl, HashSet<UInt64> excludeList)
		{
			DataTable table = new DataTable();
			table.TableName = "спецификация";

			table.Columns.Add("№ П/П", typeof(UInt64));
			table.Columns.Add("Наименование", typeof(string));
			table.Columns.Add("Пломба", typeof(UInt64));
			table.Columns.Add("КОД ТНВЭД", typeof(UInt64));
			table.Columns.Add("Кол.", typeof(UInt32));
			table.Columns.Add("Ед.изм.", typeof(string));
			table.Columns.Add("Мест", typeof(Byte));
			table.Columns.Add("УПАКОВКА", typeof(string));
			table.Columns.Add("БРУТТО", typeof(double));
			table.Columns.Add("НЕТТО", typeof(double));
			table.Columns.Add("СТОИМОСТЬ", typeof(decimal));

			var codes = codesTbl.codes;
			var groupedCodes = codesTbl.GetGroupedByCode(null);

			UInt32 itemsCount = CalcItemsCount(batchTbl);
			double packageWeight = (double)batchTbl.weightPackage / (double)itemsCount;


			Random rnd = new Random();

			var specificationData = new Dictionary<UInt32, BidloProduct>();
			var removed = new Dictionary<string, BidloProduct>();

			Common.Log("Кол-во эл-ов: " + bidloProducts.Count);

			for (int i = bidloProducts.Count - 1; i >= 0 ; i--)
			{
				var itemsWithSameName = from item in bidloProducts
										where item.name == bidloProducts[i].name
										select item;

				var itemsWithSameCode = from item in bidloProducts
										where item.code == bidloProducts[i].code
										select item;

				var itemsWithSameBagNumber = from item in bidloProducts
											 where item.bagNumber == bidloProducts[i].bagNumber
											 select item;

				if (!excludeList.Contains(bidloProducts[i].code))
				{
					if (itemsWithSameCode.ToArray().Count() > 1 && itemsWithSameBagNumber.ToArray().Count() > 1)
					{
						var tmp = bidloProducts[i];

						bidloProducts.Remove(bidloProducts[i]);

						for (var j = 0; j < bidloProducts.Count ; j++)
						{
							if (bidloProducts[j].bagNumber == tmp.bagNumber)
							{
								var item = bidloProducts[j];
								tmp.grossWeight += item.grossWeight;
								tmp.price += item.price;
								tmp.places += item.places;
								tmp.netWeight += item.netWeight;
								bidloProducts[j] = tmp;
								break;
							}
						}

					}

				}else {
					if (itemsWithSameName.ToArray().Count() > 1 && itemsWithSameBagNumber.ToArray().Count() > 1)
					{
						var tmp = bidloProducts[i];

						bidloProducts.Remove(bidloProducts[i]);

						for (var j = 0; j < bidloProducts.Count; j++)
						{
							if (bidloProducts[j].name.Equals(tmp.name))
							{
								var item = bidloProducts[j];
								tmp.grossWeight += item.grossWeight;
								tmp.price += item.price;
								tmp.places += item.places;
								tmp.netWeight += item.netWeight;

								bidloProducts[j] = tmp;
								break;
							}
						}
					}
				}
			}


			////фиксим повторения
			//Dictionary<string, BidloProduct> fixedProducts = new Dictionary<string, BidloProduct>(200);

			//foreach (var entry in removed)
			//{
			//	var item = entry.Value;

			//	BidloProduct found;
			//	var key = "" + item.code + "" + item.bagNumber;

			//	if (!fixedProducts.TryGetValue(key, out found))
			//	{
			//		fixedProducts.Add(key, item);
			//	}
			//	else {
			//		var tmp = item;

			//		tmp.grossWeight += found.grossWeight;
			//		tmp.price += found.price;
			//		tmp.places += found.places;
			//		tmp.netWeight += found.netWeight;

			//		fixedProducts[key] = tmp;
			//	}
			//}

			var bagsUsed = new HashSet<UInt64>();
			var num = 0;
			foreach (var entry in bidloProducts) {
				var item = entry;
				var bagPlace = 0;

				if (!bagsUsed.Contains(item.bagNumber))
				{
					bagPlace = 1;
					bagsUsed.Add(item.bagNumber);
				}


				++num;
				Common.Log(item.bagNumber + " " + item.name);
				//             №  П/П   Наименование  пломба     КОД ТНВЭД  Кол.   Ед.изм.  Мест        УПАКОВКА   БРУТТО       НЕТТО      СТОИМОСТЬ
				table.Rows.Add(num,              item.name,  item.bagNumber, item.code,      item.places, "шт.",   bagPlace,  "мешок",   item.grossWeight, item.netWeight, item.price);
			}

			return table;
		}
    }
}