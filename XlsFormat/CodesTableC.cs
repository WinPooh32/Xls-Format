using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public Dictionary<string, UInt64> codes = new Dictionary<string, UInt64>(1000);

		public struct ColumnNames
		{
			public string name;
			public string code;
		}

        public CodesTableC(string file, ColumnNames columnsMap)
        {
            try{
                var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);

				var enumerName = Common.getCellsEnumerator(worksheet, columnsMap.name);
				var enumberCode = Common.getCellsEnumerator(worksheet, columnsMap.code);

                //пропускаем заголовки
                enumerName.MoveNext(); enumberCode.MoveNext();

                while(enumerName.MoveNext() && enumberCode.MoveNext()){
                    string key = enumerName.Current.GetValue<string>().Trim();
                   
                    try{
						UInt64 val = Convert.ToUInt64(enumberCode.Current.GetValue<string>().Trim());
                        
						codes.Add(key, val);
                    }
                    catch(Exception ex){
                        //игнорируем повторения
                        Console.WriteLine(ex);
                    }
                }

                workbook = null;
            }
            catch(Exception ex){
                Console.WriteLine(ex);
                throw new ArgumentException("[CodesTableC] Error in file: " + file);
            }
        }
    }
}
