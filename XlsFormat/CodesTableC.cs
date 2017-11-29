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
                foreach (IXLCell cell in cellsA){
                    names.Add(cell.GetValue<String>().Trim());
                }

                //столбец 
				foreach (IXLCell cell in cellsB)
				{
                    codes.Add(cell.GetValue<String>().Trim());
				}

                workbook = null;
            }
            catch(Exception ex){
                throw new ArgumentException("Error in file: " + file);
            }
        }
    }
}
