using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
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

            List<Person> peopleFromExcel = await LoadExcelFile(file);

            foreach(var p in peopleFromExcel)
            {
                Console.WriteLine(value: $"{p.Id} {p.FirstName} {p.LastName} ");
            }
        }

        private static async Task<List<Person>> LoadExcelFile(FileInfo file)
        {
            List<Person> output = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var ws = package.Workbook.Worksheets[PositionID: 0];

            int row = 3;
            int col = 1;
            //Looks into cell determin if there is any null data/white space
            while(string.IsNullOrWhiteSpace(ws.Cells[row,col].Value?.ToString()) == false)
            {
                Person p = new();
                p.Id = int.Parse(ws.Cells[row, col].Value.ToString());
                p.FirstName = ws.Cells[row, col + 1].Value.ToString();
                p.LastName = ws.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;
            }
            return output;
        }

        private static async Task SaveExcelFile(List<Person> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add(Name:"MainReport");

            var range = ws.Cells[Address: "A2"].LoadFromCollection(people, PrintHeaders:true);

            range.AutoFitColumns();
            //Forms Header
            ws.Cells[Address: "A1"].Value = "The Report";
            ws.Cells[Address: "A1:C1"].Merge = true;
            ws.Column(col: 1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(row: 1).Style.Font.Size = 24;
            ws.Row(row: 1).Style.Font.Color.SetColor(Color.Red);

            ws.Row(row: 2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Row(row: 2).Style.Font.Bold = true;
            ws.Column(col: 3).Width = 20;
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
