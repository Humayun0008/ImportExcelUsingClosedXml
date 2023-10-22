/*using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace ImportExcelUsingClosedXml
{
    public class Program
    {
        static void Main(string[] args)
        {   //Create workboook and sheet inside it
            var wbook=new XLWorkbook();
            var wsheet=wbook.AddWorksheet("Excel_Sheet1");
            //Excel cell
            wsheet.Cell("A1").Value = "Humayun Mushtaq";
            //wsheet.ColumnWidth = 25;
            wsheet.Cell("A2").SetValue("Gujjar").SetActive(true);

            //Read the values from Excel
            var data=wsheet.Cell("A1").GetValue<string>();
            Console.WriteLine(data);    

            //Apply Style
            wsheet.Column("A").Width = 25;
            wsheet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            wsheet.Cell("A1").Style.Alignment.Vertical= XLAlignmentVerticalValues.Center;
            wsheet.Cell("A1").Style.Font.Italic = true;
            //Excel Ranges
            wsheet.Range("D2:E2").Style.Fill.BackgroundColor = XLColor.Gray;
            wsheet.Ranges("C5, F5:G8").Style.Fill.BackgroundColor = XLColor.Gray;
            var Rand = new Random();
            var Rang = wsheet.Range("H1:H5");
            wsheet.Column("H").Width = 25;
            foreach(var cell in Rang.Cells())
            {
                cell.Value = Rand.Next();
            }
            //Merge cell with ranges
            wsheet.Range("I1:J4").Merge();
            //Sort but not work
            var rand = new Random();
            var range = wsheet.Range("K1:k15");
            foreach (var cell in range.Cells())
            {
                cell.Value = rand.Next(1, 100);
            }
            wsheet.Sort("K");

            //Cell Used 
            wsheet.Cell("L1").Value = "sky";
            wsheet.Cell("L2").Value = "cloud";
            wsheet.Cell("L3").Value = "book";
            wsheet.Cell("L4").Value = "cup";
            wsheet.Cell("L5").Value = "snake";
            wsheet.Cell("L6").Value = "falcon";
            wsheet.Cell("M1").Value = "in";
            wsheet.Cell("M2").Value = "tool";
            wsheet.Cell("M3").Value = "war";
            wsheet.Cell("M4").Value = "snow";
            wsheet.Cell("M5").Value = "tree";
            wsheet.Cell("M6").Value = "ten";
            var n = wsheet.Range("L1:N10").CellsUsed().Count();
            Console.WriteLine($"There are {n} words in the range");
            //We filter all words that have three letters
            var words=wsheet.Range("L1:N10").CellsUsed()
                .Select(c=>c.Value.ToString())
                .Where(c=>c?.Length==3).ToList();
            words.ForEach(Console.WriteLine);

            wsheet.Cell("H16").FormulaA1 = "SUM(H1:H15)";
            wsheet.Cell("A8").Style.Font.Bold = true;
            wbook.SaveAs("ClosedXml.xlsx");  //Save the workbook
        }
    }
}
*/