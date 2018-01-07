using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public struct Car {
        public string name;
		public string number;
		public string numberShort;
        public string docs;
        public string vin;
    }

    public struct Driver {
        public string name;
        public string passport;
    }
    
    public class CarsTableC
    {
        public List<Car> cars = new List<Car>();
        public List<Driver> drivers = new List<Driver>();

        public CarsTableC (string file)
        {
            try{
                var workbook = new XLWorkbook(file);
                var worksheet = workbook.Worksheet(1);

                loadDrivers(worksheet);
                loadCars(worksheet);
            }
            catch(Exception ex){
                Console.WriteLine(ex);
                throw new ArgumentException("[CarsTableC] Error in file: " + file);
            }
        }
    
        private void loadDrivers(IXLWorksheet worksheet){
            var enumerName = Common.getCellsEnumerator(worksheet, "A");
            var enumerPass = Common.getCellsEnumerator(worksheet, "B");

            //пропускаем заголовки
            enumerName.MoveNext(); enumerPass.MoveNext();

            while (enumerName.MoveNext () && enumerPass.MoveNext ()) {
                drivers.Add (new Driver{
                    name = enumerName.Current.GetString(),
                    passport = enumerPass.Current.GetString()
                });
            }
        }

        private void loadCars(IXLWorksheet worksheet){
            var enumerCar = Common.getCellsEnumerator(worksheet, "D");
			var enumerNumber = Common.getCellsEnumerator(worksheet, "E");
            var enumerDocs = Common.getCellsEnumerator(worksheet, "F");
            var enumerVin = Common.getCellsEnumerator(worksheet, "G");

            //пропускаем заголовки
			enumerCar.MoveNext(); enumerNumber.MoveNext(); enumerDocs.MoveNext(); enumerVin.MoveNext();

            while (enumerCar.MoveNext() && enumerNumber.MoveNext() && enumerDocs.MoveNext() && enumerVin.MoveNext()) {
				cars.Add(new Car
				{
					name = enumerCar.Current.GetString(),
					number = enumerNumber.Current.GetString(),
					numberShort = enumerNumber.Current.GetString().Substring(1, 3),
                    docs = enumerDocs.Current.GetString(),
                    vin = enumerVin.Current.GetString()
                });
            }
        }
    }
}

