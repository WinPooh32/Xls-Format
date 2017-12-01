using System;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace XlsFormat
{
    public struct Car {
        public string name;
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
            var enumerDocs = Common.getCellsEnumerator(worksheet, "E");
            var enumerVin = Common.getCellsEnumerator(worksheet, "F");

            //пропускаем заголовки
            enumerCar.MoveNext(); enumerDocs.MoveNext(); enumerVin.MoveNext();

            while (enumerCar.MoveNext() && enumerDocs.MoveNext() && enumerVin.MoveNext()) {
                cars.Add (new Car{
                    name = enumerCar.Current.GetString(),
                    docs = enumerDocs.Current.GetString(),
                    vin = enumerVin.Current.GetString()
                });
            }
        }
    }
}

