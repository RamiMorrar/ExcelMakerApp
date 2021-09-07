using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelMaker
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(fileName:@"Demonstration.xlsx");

            var people = GetSetupData();
        }

        static List<Person> GetSetupData()
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
