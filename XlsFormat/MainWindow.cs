using System;
using System.Collections.Generic;
using Gtk;
using XlsFormat;

public partial class MainWindow : Gtk.Window
{
	private CodesTableC tableCodes;
	private BatchTableC tableBatch;
	private CarsTableC tableCars;
	//private string pathTemplatePackingList;

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
		//hpaned1.Position = 200;

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
		btnNextTransport.Clicked += OnTransportNext;

	}

	protected void Warning(string message) 
	{
		this.Sensitive = false;

		var md = new MessageDialog(this, DialogFlags.Modal | DialogFlags.DestroyWithParent, 
		                           MessageType.Warning, ButtonsType.Ok, message);
		md.Run();
		md.Destroy();

		this.Sensitive = true;
	}

	protected void fileChooserWarning(string filePath)
	{
		if (string.IsNullOrEmpty(filePath))
			{
                Warning("Файл не выбран!");
			}
			else
			{
				Warning("Ошибка в файле: '" + filePath + "'");
			}
	}

	protected void generateTittleLists(Gtk.ComboBox cb)
	{
		string defaultVal = cb.ActiveText;

		ClearCombo(cb);

		cb.AppendText(defaultVal);

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

		cb.Active = 0;
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

	protected void NextPage(object sender, EventArgs e)
	{
		if (stackPages.Page < stackPages.NPages)
		{
			stackPages.NextPage();
		}
	}

	protected void PrevPage(object sender, EventArgs e)
	{
		if ( stackPages.CurrentPage > 0)
		{
			stackPages.PrevPage();
		}	}

	protected void OnTNVDNext(object sender, EventArgs e)
	{
		var filePath = ExtractChooserPath(filechooserTNVED);

		try
		{
			tableCodes = new CodesTableC(filePath, makeCodesMap());

			NextPage(sender, e);
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			fileChooserWarning(filePath);
		}	}

	protected void OnPartyNext(object sender, EventArgs e)
	{
		var filePath = ExtractChooserPath(filechooserParty);

		try
		{
			tableBatch = new XlsFormat.BatchTableC();
			var error = tableBatch.Load(filePath, makeBatchMap());

			if (!string.IsNullOrEmpty(error))
			{
				Common.Log(error);
				Warning(error);
			}
			else
			{
				NextPage(sender, e);
			}
		}
		catch (KeyNotFoundException knfe)
		{
			Warning("Мешок с номером '" + knfe.Message + "' не найден в таблице товаров.");
		}
		catch(Exception ex)
		{
			Console.WriteLine(ex);
            fileChooserWarning(filePath);
		}
	}

	protected void OnTransportNext(object sender, EventArgs e)
	{
		NextPage(sender, e);
	}

	protected void OnTemplateNext(object sender, EventArgs e)
	{
		NextPage(sender, e);	}

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
		var filePath = ExtractChooserPath(sender);

		try
		{
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
			Console.WriteLine(ex);
			if (string.IsNullOrEmpty(filePath))
			{
                Warning("Файл транспортной БД не выбран!");
			}
			else
			{
                Warning("Ошибка в файле транспортной БД:\r\n'" + filePath + "'");
			}
		}
	}

	private void SavePackingList(string path)
	{
		try
		{
			var generatorPacking = new PackingGeneratorC();

			var carIdx = combCar.Active;
			var driverIdx = combDriver.Active;

			//Загружаем коды
			var filePath = ExtractChooserPath(filechooserTNVED);

			//Принудительно выгружаем коды из памяти===============
			tableCodes.ForceCleanup();
			tableCodes = null;
			GC.WaitForFullGCComplete();

			tableCodes = new CodesTableC(filePath, makeCodesMap());
			//=====================================================

			var driver = tableCars.drivers[driverIdx];
			var car = tableCars.cars[carIdx];

			var retCode = generatorPacking.generatePackingList(
				    ExtractChooserPath(filechooserTemplatePackingList),
					path, 
			        tableBatch, 
				    tableCodes, 
					car,
					driver,
					entryPartyNumber.Text
			);

			if (retCode == 1)
			{
				Warning("В БД кодов ТНВЭД были добавлены недостающие наименования. Пожалуйста, заполните значения и попытайтесь снова сохранить результат.");
			}
			else
			{
				generatorPacking.GenerateSpecification(ExtractChooserPath(filechooserTemplateSpecification), 
				                                       path, 
				                                       tableBatch, 
				                                       tableCodes, 
				                                       car, 
				                                       driver,
				                                       entryPartyNumber.Text);

				generatorPacking.GenerateCMR(ExtractChooserPath(filechooserTemplateCPM),
													   path,
													   tableBatch,
													   tableCodes,
													   car,
													   driver,
													   entryPartyNumber.Text,
				                             		   entrySenderCity.Text);
					
                Warning("Файлы успешно сохранены!");
			}
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
            Warning("Не удалось сохранить файлы!");
		}
	}

	protected void OnSave(object sender, EventArgs e)
	{
		Gtk.FileChooserDialog filechooser =
		new Gtk.FileChooserDialog("Выберите папку для сохранения",
		this,
        FileChooserAction.SelectFolder,
		"Омена", ResponseType.Cancel,
		"Сохранить", ResponseType.Accept);

		if (filechooser.Run() == (int)ResponseType.Accept)
		{
			SavePackingList(filechooser.Filename);
			filechooser.Destroy();
		}
		else
		{
			filechooser.Destroy();
		}
	}

	protected void OnPackSaveAs(object sender, AddedArgs args)
	{
		var folderPath = ExtractChooserPath(sender);
	}

	//private void OnTemplatePackingSelected(object sender, EventArgs e)
	//{
	//	pathTemplatePackingList = ExtractChooserPath(sender);
	//}
}