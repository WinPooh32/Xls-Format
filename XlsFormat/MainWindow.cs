using System;
using Gtk;
using XlsFormat;

public partial class MainWindow : Gtk.Window
{
	private CodesTableC tableCodes;
	private BatchTableC tableBatch;
	private CarsTableC tableCars;
	private string pathTemplatePackingList;

	public MainWindow() : base(Gtk.WindowType.Toplevel)
	{
		Build();

		//Выставляем фильтр поиска для выбора файлов
		FileFilter xlsFilter = new FileFilter();
		xlsFilter.Name = ".xlsx";
		xlsFilter.AddPattern("*.xlsx");

		filechooserParty.AddFilter(xlsFilter);
		filechooserTNVED.AddFilter(xlsFilter);
		filechooserTransport.AddFilter(xlsFilter);

		//Начальные свойства виджетов
		stackPages.Page = 0;
		hpaned1.Position = 200;
	}

	protected void ClearCombo(Gtk.ComboBox cb)
	{
        cb.Clear();
        CellRendererText cell = new CellRendererText();
		cb.PackStart(cell, false);
        cb.AddAttribute(cell, "text", 0);
        ListStore store = new ListStore(typeof(string));
		cb.Model = store;	}

	protected void OnDeleteEvent(object sender, DeleteEventArgs a)
	{
		Application.Quit();
		a.RetVal = true;
	}

	protected void OnNext(object sender, EventArgs e)
	{
		if (stackPages.Page < stackPages.NPages)
		{
			stackPages.NextPage();
		}
		else
		{
			
		}
	}

	protected void OnBack(object sender, EventArgs e)
	{
		if ( stackPages.CurrentPage > 0)
		{
			stackPages.PrevPage();
		}	}

	private string ExtractChooserPath(object chooser)
	{
		Gtk.FileChooser fileChooser = (Gtk.FileChooser)chooser;
		return fileChooser.Filename;
	}

	protected void onTNVEDselected(object sender, EventArgs e)
	{
		try
		{
			var filePath = ExtractChooserPath(sender);
			tableCodes = new XlsFormat.CodesTableC(filePath);

			frameTNVED.Sensitive = true;
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
		}
	}

	protected void OnPartySelected(object sender, EventArgs e)
	{
		try
		{
			var filePath = ExtractChooserPath(sender);
			tableBatch = new XlsFormat.BatchTableC(filePath);

			framePartyList1.Sensitive = true;
			framePartyList2.Sensitive = true;
		}
		catch(Exception ex)
		{
			Console.WriteLine(ex);
		}
	}
	protected void OnTransportSelected(object sender, EventArgs e)
	{
		try
		{
			var filePath = ExtractChooserPath(sender);
			tableCars = new XlsFormat.CarsTableC(filePath);

			ClearCombo(combDriver);
			foreach (var driver in tableCars.drivers)
			{
				combDriver.AppendText(driver.name);
			}

            ClearCombo(combCar);
			foreach (var cars in tableCars.cars)
			{
				combCar.AppendText(cars.name);
			}

			frameTransport.Sensitive = true;
		}
		catch(Exception ex)
		{
			Console.WriteLine(ex);;
		}
	}

	private void SavePackingList(string path)
	{
		try
		{
			var generatorPacking = new PackingGeneratorC(pathTemplatePackingList);

			generatorPacking.generatePackingList(
					path + "\\Упаковочный лист.xlsx", 
			        tableBatch, tableCodes, 
					tableCars.cars[combCar.Active],
					tableCars.drivers[combCar.Active],
			        "NOMER@12738"
			);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
		}

	}

	protected void OnSave(object sender, EventArgs e)
	{
		Gtk.FileChooserDialog filechooser =
		new Gtk.FileChooserDialog("Выберите папку для сохранения",
		this,
        FileChooserAction.SelectFolder,
		"Омена", ResponseType.Cancel,
		"Открыть", ResponseType.Accept);

	    if (filechooser.Run() == (int)ResponseType.Accept) 
	    {
	            SavePackingList(filechooser.Filename);
	    }

	    filechooser.Destroy();
	       
	}

	protected void OnPackSaveAs(object sender, AddedArgs args)
	{
		var folderPath = ExtractChooserPath(sender);
	}

	private void OnTemplatePackingSelected(object sender, EventArgs e)
	{
		pathTemplatePackingList = ExtractChooserPath(sender);
	}
}
