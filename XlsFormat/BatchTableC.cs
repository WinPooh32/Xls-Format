using System;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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
        }

        public struct Product  
        {  
            //номер у нас будет ключем в словаре
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

        public Dictionary<string, List<Product>> goods = new Dictionary<string, List<Product>>(1000);

        private Regex priceRegex = new Regex(@"^[\,\.\d+]*");

        public BatchTableC (string file)
        {
            try{
                var workbook = new XLWorkbook(file);

                ColumnNames columnsMap = new ColumnNames{
                    number = "A",
                    allPlaces = "T",
                    placesByType ="AG",//FIXME возможно надо поменять местами
                    name = "AK",
                    price = "AO",

                    bagOrderNumber = "A",
                    bagNumber = "B",
                    bagWeight = "C"
                };

                loadGoods(workbook.Worksheet(1), columnsMap);//Лист “Товары”
                loadBags(workbook.Worksheet(2), columnsMap);
            }
            catch(Exception ex){
                Console.WriteLine(ex);
                throw new ArgumentException("[BatchTableC] Error in file: " + file);
            }
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
                    key = enumerNumberColumn.Current.GetValue<string>().Trim();
                    val = new Product { 
                        allPlaces       = Convert.ToUInt32(enumerAllPlacesColumn.Current.GetValue<string>().Trim()),
                        placesByType    = Convert.ToUInt32(enumerPlacesByTypeColumn.Current.GetValue<string>().Trim()),
                        name            = enumerNameColumn.Current.GetValue<string>().Trim(),
                        price           = Convert.ToDecimal( normalizePrice(enumerPrice.Current.GetValue<string>().Trim()) )
                    };

                    List<Product> values;

                    if (!goods.TryGetValue(key, out values)) {
                        values = new List<Product>();
                        goods.Add(key, values);
                    }

                    values.Add(val);
                }
                catch(Exception ex){
                    //игнорируем повторения ключа
                    //TODO уведомление

                    Console.WriteLine (i);
                    Console.WriteLine (ex);
                }
            }
        }

        private void loadBags(IXLWorksheet worksheet, ColumnNames columnsMap){
            var bagOrderNumberColumn    = Common.getCellsEnumerator (worksheet, columnsMap.bagOrderNumber);
            var enumerBagNumberColumn   = Common.getCellsEnumerator (worksheet, columnsMap.bagNumber);
            var enumerBagWeightColumn   = Common.getCellsEnumerator (worksheet, columnsMap.bagWeight);

            //Пропускаем заголовки
            bagOrderNumberColumn.MoveNext();
            enumerBagNumberColumn.MoveNext();
            enumerBagWeightColumn.MoveNext();

            int i = 0;

            while (
                bagOrderNumberColumn.MoveNext()  &&
                enumerBagNumberColumn.MoveNext() &&
                enumerBagWeightColumn.MoveNext()
            ) {
                ++i;

                try{
                    var key = bagOrderNumberColumn.Current.GetValue<string>().Trim();
                    var bagNumber = Convert.ToUInt64(enumerBagNumberColumn.Current.GetString().Trim());
                    var bagWeight = Convert.ToDouble(enumerBagWeightColumn.Current.GetString().Trim());

                    var list = goods[key];

                    for(int k = 0; k < list.Count; ++k){
                        var value = list[k];

                        value.bagNumber = bagNumber;
                        value.bagWeight = bagWeight;

                        list[k] = value;
                    }
                }
                catch(Exception ex){
                    Console.WriteLine (i);
                    Console.WriteLine (ex);
                }
            }
        }

        private string normalizePrice(string rawPrice){
            Match match = priceRegex.Match(rawPrice);

            if (match.Success) {
                //для формата decimal
                //https://msdn.microsoft.com/ru-ru/library/cafs243z(v=vs.110).aspx
                return match.Value.Replace (',', '.');
            } else {
                throw new ArgumentException("[BatchTableC] Broken price: " + rawPrice);
            }
        }
    }
}

