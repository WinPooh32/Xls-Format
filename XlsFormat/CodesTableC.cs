using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public class CodesTableC
    {
        public Dictionary<string, UInt64> codes = new Dictionary<string, UInt64>(1000);

        public CodesTableC(string file)
        {
            try{
                var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);

                var enumerA = Common.getCellsEnumerator(worksheet, "A");
                var enumerB = Common.getCellsEnumerator(worksheet, "B");

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
