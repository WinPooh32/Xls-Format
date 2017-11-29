using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public Dictionary<String, UInt64> codes = new Dictionary<String, UInt64>(500);
//        public ArrayList names = new ArrayList();
//        public ArrayList codes = new ArrayList();

        public CodesTableC(String file)
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

                int i = 0;
                while(enumerA.MoveNext() && enumerB.MoveNext()){
                    var key = enumerA.Current.GetValue<String>().Trim();
                    var val = Convert.ToUInt64(enumerB.Current.GetValue<String>().Trim());

                    try{
                        codes.Add(key, val);
                    }
                    catch(Exception ex){
                    
                    }


                    key = null;

                    Console.WriteLine(++i);
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
