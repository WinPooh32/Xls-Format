using System;
using System.Collections;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public ArrayList names = new ArrayList();
        public ArrayList codes = new ArrayList();

        public CodesTableC(String file)
        {
            try{
                var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);

                var nameColumn = worksheet.Columns("A");
                var codeColumn = worksheet.Columns("B");

                var cellsA = nameColumn.CellsUsed();
                var cellsB = codeColumn.CellsUsed();

                //считываем столбец имен
                var excludeHeader = true;
                foreach (IXLCell cell in cellsA){
                    if(excludeHeader){
                        excludeHeader = false;
                        continue;
                    }
                    names.Add(cell.GetValue<String>().Trim());
                }

                //столбец кодов
                excludeHeader = true;
				foreach (IXLCell cell in cellsB)
				{
                    if(excludeHeader){
                        excludeHeader = false;
                        continue;
                    }
                    Console.WriteLine(cell.GetValue<String>().Trim());
                    codes.Add(Convert.ToUInt64(cell.GetValue<String>().Trim(), 10));
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
