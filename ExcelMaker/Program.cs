using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelMaker
{
    class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(fileName:@"Demonstration.xlsx");

            var people = GetSetupData();

            await SaveExcelFile(people, file);
        }

        private static async Task SaveExcelFile(List<Person> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add(Name:"MainReport");

            var range = ws.Cells[Address: "A1"].LoadFromCollection(people, PrintHeaders:true);

            range.AutoFitColumns();

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static List<Person> GetSetupData()
        {
            List<Person> output = new() // << new C# 9 syntax that means new Person();
            {
                new() {Id = 1 , FirstName = "Rami", LastName = "Morrar"},
                new() { Id = 1, FirstName = "Terry", LastName = "Bogard" },
                new() { Id = 1, FirstName = "GoldLewis", LastName = "Dickinson" },

            };
            return output;
        }
    }
}
