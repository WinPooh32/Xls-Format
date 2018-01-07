using System;
using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;
using Gtk;

namespace XlsFormat
{
	class Common{

		public const string fileParty = "Упаковочный лист.xlsx";
		public const string fileSpecification = "Спецификация.xlsx";
		public const string fileCMR = "СМР.xlsx";
		public const string fileNotFoundCodes = "Коды ТН ВЭД.xlsx";

		const string logsPath = "Logs";
		static TextWriter logWriter;

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

		public static void InitLogger()
		{			
	        try 
	        {
				// Determine whether the directory exists.
				if (!Directory.Exists(logsPath))
				{
					// Try to create the directory.
					DirectoryInfo di = Directory.CreateDirectory(logsPath);
				}

				string file = logsPath + "/" + DateTime.Now.ToLongTimeString() + ".log";
				StreamWriter w = File.AppendText(file.Replace(":", "-"));
				logWriter = w;
	        } 
	        catch (Exception e) 
	        {
	            Console.WriteLine("The process failed: {0}", e);
	        }
		}

		public static void Log(string logMessage)
		{
			logWriter.WriteLine("[{0}] {1}", DateTime.Now.ToLongTimeString(), logMessage);
			logWriter.Flush();		}
    }

	class MainClass
	{
		public static void Main(string[] args)
		{
			Common.InitLogger();
			Common.Log("Запуск программы");
			
			Application.Init();
			MainWindow win = new MainWindow();
			win.Show();
			Application.Run();
		}
	}
}
