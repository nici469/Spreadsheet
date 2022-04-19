using System;
using System.Collections.Generic;
using System.Text;

namespace Spreadsheet
{
    using ClosedXML.Excel;
    class ExcelClass
    {
        //IXLWorkbook 
        public void Create(String filepath)
        {
            IXLWorkbook workbook = new XLWorkbook();
            IXLWorksheet worksheet = workbook.Worksheets.Add("SampleSheet");
            worksheet.Cell(1, 1).Value = "Hello Michael";
            worksheet.Cell("A2").Value = "A2";
            worksheet.Cell(1,10).Value = 100;
            worksheet.Cell("A2").Style.Font.Bold=true;
            worksheet.Cell("A2").Style.Font.Italic = true;
            //worksheet.Cell("A2").Style.Border.SetOutsideBorder(,) = true
            workbook.SaveAs(filepath);
        }


        public void Create2(String filePath)
        {
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sample Sheet",1);
            wb.Worksheets.Add("Export", 2);
            
            //wb.Worksheet("Export").Position = 1;

            

            //wb.Worksheet("Sample Sheet").Position = 3;

            ws.Cell(2, 3).Value = "Hello World!";
            ws.Cell(2, 3).Style.Font.Bold = true;
            ws.Cell(2, 3).Style.Font.Italic = true;


            ws.Cell(4, 2).Value = "Project:";
            ws.Cell(4, 4).Value = "ClosedXML Example";
            ws.Cell(6, 2).Value = "Author:";
            ws.Cell(6, 4).Value = "KnapSac";

            ws.Cell(2, 3).Style.Fill.SetBackgroundColor(XLColor.Cyan);

            IXLRange range = ws.Range(ws.Cell(4, 2).Address, ws.Cell(6, 4).Address);

            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            range.Style.Font.FontColor = XLColor.Blue;

            wb.Worksheet(1).FirstRowUsed().Style.Fill.BackgroundColor = XLColor.Green;

            wb.SaveAs(filePath);
        }

        /// <summary>
        /// creates a spreadsheet from a single string made up of several comma-separatedvalue lines
        /// </summary>
        /// <param name="CSVData"></param>
        /// <param name="filepath"></param>
        public void CreateXlFromCSV(string CSVData, string filepath)
        {
            var processor = new ProcessString();
            var dataJGD = processor.CSVtoJGD(CSVData);

            //init Spreadsheet
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1", 1);

            //loop through the rows in the csv data jgd array
            for(int i = 0; i < dataJGD.Length; i++)
            {
                string[] lineData = dataJGD[i];

                //loop through the columns in the lineData from csv data jgd array
                for(int j = 0; j < lineData.Length; j++)
                {
                    //info cell occurs every 2 cells in each row, starting from the second cell at j=1 index
                    if ((j + 1) % 2 == 0) { ImplementCellInfo(); continue; }

                    //if a data cell is found at the end of the line and it has no info cell after it or there is 
                    //no other cell after it, ignore the cell and break out of column iteration[j],...
                    //such cells occur because of how the csv is create in javascript
                    if (j + 1 >= lineData.Length) { break; }
                    ws.Cell(i + 1, (j + 2) / 2).Value = double.Parse(lineData[j]);
                }
            }

            wb.SaveAs(filepath);
        }


        /// <summary>
        /// Creates a spreadsheet from an array of CSV strings
        /// </summary>
        /// <param name="CSVData"></param>
        /// <param name="filepath"></param>
        public void CreateXlFromCSV(string[] CSVData, string filepath)
        {
            var processor = new ProcessString();
            var dataJGD = processor.StringArraytoJGD(CSVData);

            //init Spreadsheet
            var wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Sheet1", 1);

            //loop through the rows in the csv data jgd array
            for (int i = 0; i < dataJGD.Length; i++)
            {
                string[] lineData = dataJGD[i];

                //loop through the columns in the lineData from csv data jgd array
                for (int j = 0; j < lineData.Length; j++)
                {
                    //info cell occurs every 2 cells in each row, starting from the second cell at j=1 index
                    if ((j + 1) % 2 == 0) { ImplementCellInfo(); continue; }

                    //if a data cell is found at the end of the line and it has no info cell after it or there is 
                    //no other cell after it, ignore the cell and break out of column iteration[j],...
                    //such cells occur because of how the csv is create in javascript
                    if (j + 1 >= lineData.Length) { break; }
                    //ignore any empty data cells

                    if (lineData[j] != null) {
                        //attempt to convert the data to a double
                        try {
                            ws.Cell(i + 1, (j + 2) / 2).Value = double.Parse(lineData[j]);
                        }
                        //if the data cannot be converted to a double, store it as a string
                        catch (Exception e) {
                            ws.Cell(i + 1, (j + 2) / 2).Value = lineData[j];
                        }
                         
                    }
                    
                }
            }

            wb.SaveAs(filepath);
        }


        void ImplementCellInfo() { }

    }
}
