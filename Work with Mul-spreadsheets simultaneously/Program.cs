using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace Work_with_Mul_spreadsheets_simultaneously
{
    class Program
    {
        static void Main(string[] args)
        {
            using (SLDocument sl = new SLDocument())
            {
                SLDocument firstdoc = new SLDocument("MultipleSpreadsheetFirst.xlsx", "Sheet1");
                SLDocument seconddoc = new SLDocument("MultipleSpreadsheetSecond.xlsx", "Sheet2");

                sl.SetCellValue(4, 2, "Things to bring");

                sl.SetCellValue(5, 2, firstdoc.GetCellValueAsString("B2"));

                sl.SetCellValue(6, 2, seconddoc.GetCellValueAsString("B2"));

                sl.SetCellValue(7, 2, "Party hats");

                SLStyle style = firstdoc.GetCellStyle("B2");
                sl.SetCellStyle(6, 2, style);

                // because apparently the style we want is on another sheet...
                seconddoc.SelectWorksheet("Sheet3");
                style = seconddoc.GetCellStyle("B4");
                sl.SetCellStyle(7, 2, style);

                // get the style again. Because I'm inefficient...
                style = firstdoc.GetCellStyle("B2");
                style.SetFontUnderline(UnderlineValues.Single);
                //style.SetPatternFill(SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Accent5Color, SLThemeColorIndexValues.Accent5Color);
                //第一个参数是指填充的样式，第二和第三个参数分别是前景和背景，最后叠加成填充的颜色
                style.SetPatternFill(DocumentFormat.OpenXml.Spreadsheet.PatternValues.DarkGray  , System.Drawing.Color.Blue , System.Drawing.Color.Silver );
                firstdoc.CloseWithoutSaving();

                seconddoc.SelectWorksheet("Sheet1");
                seconddoc.SetCellStyle(5, 2, style);
                seconddoc.SetCellValue(5, 2, "Remember to bring crackers too!");

                seconddoc.SaveAs("MultipleSpreadsheetSecondModified.xlsx");

                sl.SaveAs("MultipleSpreadsheet.xlsx");
            }

            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
