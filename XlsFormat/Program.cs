using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using Gtk;

namespace XlsFormat
{
	class Common{
        private Common(){
        }

        static public IEnumerator<IXLCell> getCellsEnumerator(IXLWorksheet worksheet, string column){
            return worksheet.Column(column).CellsUsed().GetEnumerator();
        }

        static public string getCellString (IXLWorksheet ws, string cell){
            return ws.Cell (cell).GetString();
        }

        static public void setCellString (IXLWorksheet ws, string cell, string value){
            ws.Cell (cell).SetValue (value);
        }
    }

	class MainClass
	{
		public static void Main(string[] args)
		{
			Application.Init();
			MainWindow win = new MainWindow();
			win.Show();
			Application.Run();
		}
	}
}
