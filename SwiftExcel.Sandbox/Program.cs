using System;
using SwiftExcel.Extensions;
using System.Collections.Generic;

namespace SwiftExcel.Sandbox
{
    internal class Program
    {
        private const string FilePath = "C:/temp/test.xlsx";

        private static ExcelWriter _excelWriter;

        private static void Main()
        {
            var MAX_SHEETS = 2;
            var MAX_ROWS = 100;
            var MAX_COLS = 50;
            
            var programDt = DateTime.Now;
            Console.WriteLine($"Writing {MAX_COLS}x{MAX_ROWS}, {MAX_SHEETS}sheets to [{FilePath}]");
            Console.WriteLine($"START AT: {programDt}");

            using (_excelWriter = new ExcelWriter(FilePath))
            {
                for (int i = 1; i <= MAX_SHEETS; i++)
                {   
                    var sheet = _excelWriter.CreateSheet("Sheet " + i);
                    
                    // 1. Data
                    var dt = DateTime.Now;
                    Console.WriteLine($"    {sheet.Name} - DATA START AT: {dt}");
                    
                    for (int row = 1; row <= MAX_ROWS; row++)
                    {
                        for (int col = 1; col <= MAX_COLS; col++)
                        {
                            _excelWriter.Write(sheet.Name, $"row:{row}-col:{col}", col, row);
                        }
                    }
                    
                    var sheetDoneDt = DateTime.Now;
                    Console.WriteLine($"    {sheetDoneDt} - DATA DONE AT: ({(sheetDoneDt - dt).TotalMilliseconds})");
                    Console.WriteLine();
                }
                
                _excelWriter.Save();
            }
            
            // //Set custom sheet name, define columns width, right to left and wrap text
            // //Use manual Save() instead of using block 
            // var sheet = new Sheet
            // {
            //     Name = "Monthly Report",
            //     RightToLeft = false,
            //     WrapText = true,
            //     ColumnsWidth = new List<double> { 10, 12, 8, 8, 35 }
            // };
            //
            // _excelWriter = new ExcelWriter(FilePath, sheet);
            // for (var row = 1; row <= 100; row++)
            // {
            //     for (var col = 1; col <= 10; col++)
            //     {
            //         _excelWriter.Write($"row:{row}-col:{col}", col, row);
            //     }
            // }
            //
            // _excelWriter.Save();

            //
            // //Formula examples
            // using (_excelWriter = new ExcelWriter(FilePath))
            // {
            //     const int col = 1;
            //     var row = 1;
            //     for (; row <= 20; row++)
            //     {
            //         _excelWriter.Write(row.ToString(), col, row, DataType.Number);
            //     }
            //
            //     _excelWriter.WriteFormula(FormulaType.Average, col, ++row, col, 1, 20);
            //     _excelWriter.WriteFormula(FormulaType.Count, col, ++row, col, 1, 20);
            //     _excelWriter.WriteFormula(FormulaType.Max, col, ++row, col, 1, 20);
            //     _excelWriter.WriteFormula(FormulaType.Sum, col, ++row, col, 1, 20);
            // }
            //
            //
            // //Initiate test collection
            // var testCollection = new List<TestModel>
            // {
            //     new TestModel(), new TestModel()
            // };
            //
            // //Export list of objects to Excel file
            // testCollection.ExportToExcel(FilePath);
            //
            //
            // //Export list of objects to Excel file with predefined Sheet name
            // testCollection.ExportToExcel(FilePath, sheetName: "Sheet2");

            var programDoneDt = DateTime.Now;
            Console.WriteLine($"DONE AT: {programDoneDt}({(programDoneDt - programDt).TotalMilliseconds})");
        }
    }
}