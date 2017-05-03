using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Spire.Pdf;
using Spire.Xls.Converter;
using System;
using System.IO;
using System.Linq;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting process...");
            var workbook = new XLWorkbook(@"teste.xlsx");

            var dataSheet = workbook.Worksheet("Data");
            var viewSheet = workbook.Worksheet("View");

            var rnd = new Random();

            for (int i = 1; i <= dataSheet.Rows().Count(); i++)
            {
                for (int j = 1; j <= 3; j++)
                {
                    var data = dataSheet.Row(i).Cell(j).Value;
                    if (string.IsNullOrEmpty(data.ToString()))
                        dataSheet.Row(i).Cell(j).Value = rnd.Next(10, 100);
                }
            }

            workbook.SaveAs("sample.xlsx");

            Console.WriteLine("Complete. Press any key to leave.");
            Console.ReadLine();
        }
    }
}
