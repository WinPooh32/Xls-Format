using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public Dictionary<string, UInt64> codes = new Dictionary<string, UInt64>(500);

        public CodesTableC(string file)
        {
            try{
                var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);

                var nameColumn = worksheet.Columns("A");
                var codeColumn = worksheet.Columns("B");
               
                var cellsA = nameColumn.CellsUsed();
                var cellsB = codeColumn.CellsUsed();

                var enumerA = cellsA.GetEnumerator();
                var enumerB = cellsB.GetEnumerator();

                //пропускаем заголовки
                enumerA.MoveNext(); enumerB.MoveNext();

                while(enumerA.MoveNext() && enumerB.MoveNext()){
                    string key = enumerA.Current.GetValue<string>().Trim();
                    UInt64 val = Convert.ToUInt64(enumerB.Current.GetValue<string>().Trim());

                    try{
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
