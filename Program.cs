using System;

namespace Spreadsheet
{
    using IronXL;
    class Program
    {
      
        
    static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            WorkBook workbook = WorkBook.Create(ExcelFileFormat.XLSX);
            var sheet = workbook.CreateWorkSheet("example_sheet");
            sheet["A1"].Value = "Example";

            //set value to multiple cells
            sheet["A2:A4"].Value = 5;
            sheet["A5"].Style.SetBackgroundColor("#f0f0f0");

            //set style to multiple cells
            sheet["A5:A6"].Style.Font.Bold = true;

            //set formula
            sheet["A6"].Value = "=SUM(A2:A4)";
            if (sheet["A6"].IntValue == sheet["A2:A4"].IntValue)
            {
                Console.WriteLine("Basic test passed");
            }
            workbook.SaveAs("example_workbook.xlsx");

            Console.WriteLine(args.Length);

        }
    }
}
