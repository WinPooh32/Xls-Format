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

        ////////
		generateTittleLists(combTNVEDname);
		generateTittleLists(combTNVEDcode);

		generateTittleLists(combBagOrderNumber);
        generateTittleLists(combBagNumber);
		generateTittleLists(combWeight);

        generateTittleLists(combPartyName);
		generateTittleLists(combPartyNum);
		generateTittleLists(combPartyCost);
		generateTittleLists(combPartyAllCount);
		generateTittleLists(combPartyCountByType);

		btnNextTNVED.Clicked += OnTNVDNext;
		btnNextParty.Clicked += OnPartyNext;
	}

	protected void generateTittleLists(Gtk.ComboBox cb)
	{
		ClearCombo(cb);

		for (char c = 'A'; c <= 'Z'; c++)
		{
			cb.AppendText(""+c);
		}

		for (char c = 'A'; c <= 'Z'; c++)
		{
			for (char d = 'A'; d <= 'Z'; d++)
			{
				cb.AppendText(""+c+d);
			}
		}
	}

	protected void ClearCombo(Gtk.ComboBox cb)
	{
        cb.Clear();
        CellRendererText cell = new CellRendererText();
		cb.PackStart(cell, false);
        cb.AddAttribute(cell, "text", 0);
        ListStore store = new ListStore(typeof(string));
		cb.Model = store;	}

	protected CodesTableC.ColumnNames makeCodesMap()
	{
		return new CodesTableC.ColumnNames
		{
			name = combTNVEDname.ActiveText,
			code = combTNVEDcode.ActiveText
		};
	}

	protected BatchTableC.ColumnNames makeBatchMap()
	{
		return new BatchTableC.ColumnNames
		{
			allPlaces = combPartyAllCount.ActiveText,

			bagNumber = combBagNumber.ActiveText,
			bagOrderNumber = combBagOrderNumber.ActiveText,
			bagWeight = combWeight.ActiveText,

			name = combPartyName.ActiveText,
			number = combPartyNum.ActiveText,
			placesByType = combPartyCountByType.ActiveText,
			price = combPartyCost.ActiveText,

			sumGrossWeight = entrySumGross.Text,
			sumNetWeight = entrySumNetWeight.Text,
			sumPackagesWeight = entrySumPackageWeight.Text
		};
	}

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
	}

	protected void OnBack(object sender, EventArgs e)
	{
		if ( stackPages.CurrentPage > 0)
		{
			stackPages.PrevPage();
		}	}

	protected void OnTNVDNext(object sender, EventArgs e)
	{
		try
		{
			var filePath = ExtractChooserPath(filechooserTNVED);
			tableCodes = new CodesTableC(filePath, makeCodesMap());
			Console.WriteLine("OK");
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
		}	}

	protected void OnPartyNext(object sender, EventArgs e)
	{
		try
		{
			var filePath = ExtractChooserPath(filechooserParty);
			tableBatch = new XlsFormat.BatchTableC(filePath, makeBatchMap());
			Console.WriteLine("OK");
		}
		catch(Exception ex)
		{
			Console.WriteLine(ex);
		}
	}


	private string ExtractChooserPath(object chooser)
	{
		Gtk.FileChooser fileChooser = (Gtk.FileChooser)chooser;
		return fileChooser.Filename;
	}

	protected void onTNVEDselected(object sender, EventArgs e)
	{
		frameTNVED.Sensitive = true;
	}

	protected void OnPartySelected(object sender, EventArgs e)
	{
		framePartyList1.Sensitive = true;
		framePartyList2.Sensitive = true;
		framePartyList3.Sensitive = true;
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

			var savePath = path + "\\Упаковочный лист.xlsx";

			var carIdx = combCar.Active;
			var driverIdx = combDriver.Active;

			generatorPacking.generatePackingList(
					savePath, 
			        tableBatch, tableCodes, 
					tableCars.cars[carIdx],
					tableCars.drivers[driverIdx],
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
