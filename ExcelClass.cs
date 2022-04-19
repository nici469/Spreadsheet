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

    }
}
