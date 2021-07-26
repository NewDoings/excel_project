using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    partial class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(fileName: @"E:\Demos\ExelFile.xlsx");  // путь где должен создоваться файл
            var people = GetSetupData();
            await SaveExceFile(people, file);
        }

        private static async Task SaveExceFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var ws = package.Workbook.Worksheets.Add(Name: "MainReport");

            var range = ws.Cells[Address: "A2"].LoadFromCollection(people,PrintHeaders: true);
            range.AutoFitColumns();

            ws.Row(row: 1).Style.Font.Size = 24;
            ws.Cells[Address: "A2:C2"].Merge = true;
            ws.Cells[Address: "C5"].MoveNext();

            await package.SaveAsync();
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }   
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new List<PersonModel>()
            {
                new PersonModel{Id = 1, FirstName = "Tim", LastName = "Corey"},
                new PersonModel{Id = 2, FirstName = "Sue", LastName = "Storm"},
                new PersonModel{Id = 3, FirstName = "Jane", LastName = "Smith"}
            };
            return output;
        }
    }
}
