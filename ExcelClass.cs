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
                    if ((j + 1) % 2 == 0)
                    {
                        int row = i + 1;//the row number of the cell to be formatted with the info data
                        int col = (j + 1) / 2;//the column number of the cell to be formatted with the  info data
                        ImplementCellInfo(row, col, ws, lineData[j]);
                        continue;
                    }

                    //if a data cell is found at the end of the line and it has no info cell after it or there is 
                    //no other cell after it, ignore the cell and break out of column iteration[j],...
                    //such cells occur because of how the csv is create in javascript
                    if (j + 1 >= lineData.Length) { break; }
                    
                    //ignore any empty data cells
                    if (lineData[j] != null)
                    {
                        //attempt to convert the data to a double
                        try
                        {
                            ws.Cell(i + 1, (j + 2) / 2).Value = double.Parse(lineData[j]);
                        }
                        //if the data cannot be converted to a double, store it as a string
                        catch (Exception e)
                        {
                            ws.Cell(i + 1, (j + 2) / 2).Value = lineData[j];
                        }

                    }
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
            ws.Columns().Width = 9;
            ws.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);//left alignment for the worksheet

            //loop through the rows in the csv data jgd array
            for (int i = 0; i < dataJGD.Length; i++)
            {
                string[] lineData = dataJGD[i];

                //loop through the columns in the lineData from csv data jgd array
                for (int j = 0; j < lineData.Length; j++)
                {
                    //info cell occurs every 2 cells in each row, starting from the second cell at j=1 index
                    if ((j + 1) % 2 == 0) {
                        int row = i + 1;//the row number of the cell to be formatted with the info data
                        int col = (j + 1) / 2;//the column number of the cell to be formatted with the  info data

                        if (lineData[j] != null) { ImplementCellInfo(row, col, ws, lineData[j]); }
                        continue; 
                    }

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

        /// <summary>
        /// formats the specified row and column of the IXLworksheet with the infoString
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="ws"></param>
        /// <param name="infoString"></param>
        void ImplementCellInfo(int row, int col, IXLWorksheet ws, string infoString) {
            //info data in an info string are separated by hyphen character
            var infoArray = (new ProcessString()).SeparateLines(infoString, '-');

            for(int i = 0; i < infoArray.Length; i++)
            {
                string singleInfo = infoArray[i];
                //check the first character: "format" strings begin with F, e.g FHeader,
                //while color strings begin with C, e.g CBlue
                var infoType = singleInfo[0];

                switch (infoType)
                {
                    case 'C'://if the first character is a 'C"
                        switch (singleInfo) {
                            case "CRed":
                                ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Red;
                                break;
                            case "CBlue":
                                ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Blue;
                                break;
                            case "CGreen":
                                ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Green;
                                break;
                        }
                        break;

                    case 'F'://if the first character is an 'F'
                        switch (singleInfo)
                        {
                            //header only needs to be specified on one cell of the row
                            //it automatically affects the whole row
                            case "FHeader":
                                ws.Row(row).Style.Font.Italic = true;
                                ws.Row(row).Style.Font.FontSize = 13;
                                ws.Row(row).Style.Font.Bold = true;
                                //ws.Cell(row, col).Style.Font.Italic = true;
                                break;
                            case "FOutline"://for cells like the o-Clock cells at the beginning of each row
                                ws.Cell(row, col).Style.Font.Bold = true;
                                ws.Cell(row, col).Style.Font.Italic = true;
                                break;
                            case "FTitle"://for cells/rows that carry the new day title
                                IXLRange range = ws.Range(row, col, row, col + 17);
                                range.Row(1).Merge();
                                range.Style.Font.FontSize = 14;
                                range.Style.Font.Bold = true;
                                range.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                                //range.Columns().Cells().w
                                break;
                        }
                        break;

                }
            }
        }

    }
}
