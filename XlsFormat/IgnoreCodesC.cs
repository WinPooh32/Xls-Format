using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
	public class IgnoreCodesC
	{
		XLWorkbook workbook;
		IXLWorksheet worksheet;

		HashSet<UInt64> codesSet = new HashSet<UInt64>();

		public IgnoreCodesC(string file)
		{
			try
			{
				workbook = new XLWorkbook(file);
				worksheet = workbook.Worksheet(1);
				loadCodes();
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex);
				throw new ArgumentException("[IgnoreCodesC] Error in file: " + file);
			}
		}

		private void loadCodes()
		{
			var enumerCode = Common.getCellsEnumerator(worksheet, "A");

			while (enumerCode.MoveNext())
			{
				codesSet.Add(enumerCode.Current.GetValue<UInt64>());
			}
		}

		public HashSet<UInt64> GetCodes()
		{
			return codesSet;
		}
	}
}
