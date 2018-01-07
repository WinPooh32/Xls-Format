using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Linq;

namespace XlsFormat
{
    public class BatchTableC
    {
        public struct ColumnNames{
            public string number;
            public string allPlaces;
            public string placesByType;
            public string name;
            public string price;

            public string bagOrderNumber;
            public string bagNumber;
            public string bagWeight;

			public string sumNetWeight;
			public string sumGrossWeight;
			public string sumPackagesWeight;
        }

        public struct Product  
        {
			//номер у нас будет ключем в словаре
			public string number;
            public UInt32 allPlaces; //2. Суммарное количество вложений (количество мест в пакете)
            public UInt32 placesByType; // 3. Количество объектов конкретного типа в посылке
            public string name;
            public decimal price;

            //мешки
            public UInt64 bagNumber;
            public double bagWeight;
        }  

        public double weightNet;//нетто
        public double weightPackage;//вес упаковки
        public double weightGross;//брутто

		public IOrderedEnumerable<Product> sortedProducts;

        private Dictionary<string, List<Product>> goods = new Dictionary<string, List<Product>>(1000);

        private Regex priceRegex = new Regex(@"^[\,\.\d+]*");

        public BatchTableC ()
        {

        }

		public string Load(string file, ColumnNames columnsMap)
		{
			try
			{
				var workbook = new XLWorkbook(file);

				loadGoods(workbook.Worksheet(1), columnsMap);//Лист “Товары”
				var error = loadBags(workbook.Worksheet(2), columnsMap);

				if (!string.IsNullOrEmpty(error))
				{
					return error;
				}

				sortedProducts = toLinearList();
			}
			catch (KeyNotFoundException knfe)
			{
				throw knfe;
			}
			catch (Exception ex){
				Common.Log(ex.ToString());
                throw new ArgumentException("[BatchTableC] Error in file: " + file);
            }

			return null;
		}

        private void loadGoods(IXLWorksheet worksheet, ColumnNames columnsMap){
            var enumerNumberColumn          = Common.getCellsEnumerator (worksheet, columnsMap.number);
            var enumerAllPlacesColumn       = Common.getCellsEnumerator (worksheet, columnsMap.allPlaces);
            var enumerPlacesByTypeColumn    = Common.getCellsEnumerator(worksheet, columnsMap.placesByType);
            var enumerNameColumn            = Common.getCellsEnumerator (worksheet, columnsMap.name);
            var enumerPrice                 = Common.getCellsEnumerator(worksheet, columnsMap.price);

            string key;
            Product val;

            //Пропускаем заголовки
            enumerNumberColumn.MoveNext();
            enumerAllPlacesColumn.MoveNext();
            enumerPlacesByTypeColumn.MoveNext();
            enumerNameColumn.MoveNext();
            enumerPrice.MoveNext();

            int i = 0;

            while(
                enumerNumberColumn.MoveNext()       && 
                enumerAllPlacesColumn.MoveNext()    && 
                enumerPlacesByTypeColumn.MoveNext() && 
                enumerNameColumn.MoveNext()         && 
                enumerPrice.MoveNext()
            ){
                ++i;

                try{
					key = enumerNumberColumn.Current.GetValue<string>().Trim().ToLower();

                    val = new Product { 
                        allPlaces       = Convert.ToUInt32(enumerAllPlacesColumn.Current.GetValue<string>().Trim()),
                        placesByType    = Convert.ToUInt32(enumerPlacesByTypeColumn.Current.GetValue<string>().Trim()),
                        name            = enumerNameColumn.Current.GetValue<string>().Trim(),
						price           = Decimal.Parse(normalizePrice(enumerPrice.Current.GetValue<string>().Trim()), System.Globalization.NumberStyles.Number)
                    };

                    List<Product> values;

                    if (!goods.TryGetValue(key, out values)) {
                        values = new List<Product>();
                        goods.Add(key, values);
                    }

                    values.Add(val);
                }
                catch(Exception ex){
                    Console.WriteLine (ex);
					Common.Log(ex.ToString());
                }
            }
        }

        private string loadBags(IXLWorksheet worksheet, ColumnNames columnsMap){

			//считываем суммы
			weightNet = Convert.ToDouble(worksheet.Cell(columnsMap.sumNetWeight).GetString().Trim());
			weightGross = Convert.ToDouble(worksheet.Cell(columnsMap.sumGrossWeight).GetString().Trim());
			weightPackage = Convert.ToDouble(worksheet.Cell(columnsMap.sumPackagesWeight).GetString().Trim());


            var bagOrderNumberColumn    = Common.getCellsEnumerator (worksheet, columnsMap.bagOrderNumber);
            var enumerBagNumberColumn   = Common.getCellsEnumerator (worksheet, columnsMap.bagNumber);
            var enumerBagWeightColumn   = Common.getCellsEnumerator (worksheet, columnsMap.bagWeight);

            //Пропускаем заголовки
            bagOrderNumberColumn.MoveNext();
            enumerBagNumberColumn.MoveNext();
            enumerBagWeightColumn.MoveNext();

            int i = 0;

			double testSum = 0;

            while (
                bagOrderNumberColumn.MoveNext()  &&
                enumerBagNumberColumn.MoveNext() &&
                enumerBagWeightColumn.MoveNext()
            ) {
                ++i;

				try
				{
					string key = bagOrderNumberColumn.Current.GetValue<string>().Trim().ToLower();
					UInt64 bagNumber = Convert.ToUInt64(enumerBagNumberColumn.Current.GetString().Trim());
					double bagWeight = Convert.ToDouble(normalizeFloat(enumerBagWeightColumn.Current.GetString().Trim()));

					List<Product> list;

					if (!goods.TryGetValue(key, out list))
					{
						throw new KeyNotFoundException(key);
					}

					uint sumPlaces = 0;
					uint bagPlaces = 0;

					testSum += bagWeight;

					for (int k = 0; k < list.Count; ++k)
					{
						var value = list[k];

						value.bagNumber = bagNumber;
						value.bagWeight = bagWeight;

						bagPlaces = value.allPlaces;

						sumPlaces += value.placesByType;

						list[k] = value;
					}

					if (sumPlaces > bagPlaces)
					{
						//error
						return "Заказ с номером " + key + " cодержит " + sumPlaces + " вещей из " + bagPlaces + " возможных.";
					}
				}
				catch (KeyNotFoundException knfe)
				{
					Common.Log("Мешок с номером '" + knfe.Message + "' не найден");
					throw knfe;
				}
				catch (Exception ex){
                    Console.WriteLine (ex);
					Common.Log(ex.ToString());
                }
            }

			Common.Log("Сумма мешков нетто: " + testSum);

			return null;
        }

        private string normalizePrice(string rawPrice){
            Match match = priceRegex.Match(rawPrice);

            if (match.Success) {
				//для формата decimal
				//https://msdn.microsoft.com/ru-ru/library/cafs243z(v=vs.110).aspx
				var norm = match.Value.Replace('.', ',');
				Console.WriteLine(norm);
                return norm;
            } else {
                throw new ArgumentException("[BatchTableC] Broken price: " + rawPrice);
            }
        }

		private string normalizeFloat(string rawFloat)
		{
			return rawFloat.Replace('.', ',');
		}

		private IOrderedEnumerable<Product> toLinearList()
		{
			var list = new List<Product>();

			foreach (KeyValuePair<string, List<XlsFormat.BatchTableC.Product>> entry in goods)
			{
				var key = entry.Key;
				foreach (XlsFormat.BatchTableC.Product value in entry.Value)
				{
					var copy = value;
					copy.number = key;

					list.Add(copy);
				}
			}
			
			return (
				from product in list
				orderby product.bagNumber
			    select product
			);
		}

		public Dictionary<UInt64, List<Product>> GetGroupedByCode(Dictionary<string, UInt64> codes)
		{
			var grouped = new Dictionary<UInt64, List<Product>>(1000);

			foreach (Product item in sortedProducts)
			{
				UInt64 code;
				if (!codes.TryGetValue(item.name, out code))
				{
					Common.Log("Не найден код для '" + item.name + "'");
					continue;
				}

				List<Product> products;
				if (!grouped.TryGetValue(code, out products))
				{
					products = new List<Product>();
					grouped.Add(code, products);
				}

				products.Add(item);
			}

			return grouped;
		}
    }
}
