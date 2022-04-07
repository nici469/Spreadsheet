using System;

namespace Spreadsheet
{
    using IronXL;
    using System.IO;
    class Program
    {
      
        
    static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            Console.WriteLine(args.Length);
            
            //TestProcessString();
            CSVtoJGD();
        }

        static void TestSpreadsheet()
        {
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

            
        }
        static void TestProcessString()
        {
            ProcessString processor = new ProcessString();
            var strings = "1 This is a-2 string-3 separated with spaces-4 but majorly-5  by-6 hyphen";
            var lines = processor.SeparateLines(strings, '-');

            foreach (string s in lines) Console.WriteLine(s);
            Console.WriteLine("The number of separate strings is " + lines.Length);
            Console.ReadKey(true);
        }

        /// <summary>
        /// to test the conversion of a CSV file to a jagged array
        /// </summary>
        /// <param name="csvString"></param>
        static void CSVtoJGD()
        {
            string testCSV = "this,is,the,first,line\nThis,is,the,second,line\nAnd,this,then,is,the,third,line you know";
            ProcessString processor = new ProcessString();

            //separate the string into lines.. unnecessary if File.ReadAllLines is used
            string[] lines = processor.SeparateLines(testCSV, '\n');
            
            //initialise the jagged array[row][column]. the number of rows is the number of elements in the lines array
            string[][] csvJgdArray = new string[lines.Length][];

            //in each line, separtae the strings usinng commas
            for(int i = 0; i < lines.Length; i++)
            {
                csvJgdArray[i] = processor.SeparateLines(lines[i], ',');
            }

            Console.WriteLine("there are {0} lines in the given string", lines.Length);
            Console.WriteLine("the number of elements in line 3 is: {0}", csvJgdArray[2].Length);
            Console.ReadKey(true);
        }

    }
}
