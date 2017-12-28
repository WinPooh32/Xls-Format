using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public Dictionary<string, UInt64> codes = new Dictionary<string, UInt64>(1000);

		XLWorkbook workbook;
		IXLWorksheet worksheet;

		public struct ColumnNames
		{
			public string name;
			public string code;
		}

		public CodesTableC(string file, ColumnNames columnsMap)
		{
			try
			{
				workbook = new XLWorkbook(file);
				worksheet = workbook.Worksheet(1);

				var enumerName = Common.getCellsEnumerator(worksheet, columnsMap.name);
				var enumberCode = Common.getCellsEnumerator(worksheet, columnsMap.code);

				//пропускаем заголовки
				enumerName.MoveNext(); enumberCode.MoveNext();

				while (enumerName.MoveNext() && enumberCode.MoveNext())
				{
					string key = enumerName.Current.GetValue<string>().Trim();

					UInt64 val = Convert.ToUInt64(enumberCode.Current.GetValue<string>().Trim());

					UInt64 testVal;
					if (!codes.TryGetValue(key, out testVal))
					{
						codes.Add(key, val);
					}
					else
					{
						Common.Log("CodesTableC() Повтор наименования:'" + key + "'");
					}
				}
			}
			catch (Exception ex)
			{
				Common.Log("CodesTableC() ошибка при обработке файла: '"+ file + "'\r\n" + ex);
				throw new ArgumentException("[CodesTableC] Error in file: " + file);
			}
		}

		public void AppendNotFoundNames(HashSet<string> names)
		{
			var insertFrom = calcUsedRows();

			foreach (string name in names) {
				var row = worksheet.Row(++insertFrom);
				row.Cell(1).SetValue(name);
				row.Cell(2).SetValue(0);
			}

			workbook.Save();
		}

		public void ForceCleanup()
		{
			worksheet = null;
			workbook = null;
			GC.WaitForFullGCComplete();
		}

		private int calcUsedRows()
		{
			var enumer = worksheet.RowsUsed(false).GetEnumerator();
			int count = 0;
			while (enumer.MoveNext()) { ++count;}
			return count;
		}
    }
}
