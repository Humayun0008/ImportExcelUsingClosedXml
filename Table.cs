using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ImportExcelUsingClosedXml
{
     public class Table
    {
        static void Main(string[] args)
        {
           /* var wb = new XLWorkbook();
            var ws = wb.AddWorksheet()*/;
            /*            ws.ColumnWidth = 12;
                        ws.FirstCell().InsertTable(new[]
                        {
                            new Pastry("Pie", 10),
                            new Pastry("Cake", 7),
                            new Pastry("Waffles", 17)
                        }, "PastrySales", true);

                        ws.Range("D2:D5").CreateTable("Table");*/

            /* ws.Cell(1, 1).Value = "Foo";
             ws.Cell(1, 2).Value = "Bar";
             ws.Range("A2:B5").Value = 10;
             var table = ws.Range("A1:B5").CreateTable();

             table.Resize("B1", "D3");*/


            /*            ws.Cell("A1").SetValue("First");
                        ws.Cell("A2").InsertData(Enumerable.Range(1, 5));
                        ws.Cell("B1").SetValue("Second");
                        ws.Cell("B2").InsertData(Enumerable.Range(1, 5));

                        var table = ws.Range("A1:B6").CreateTable();
                        table.Theme = XLTableTheme.TableStyleLight16;

                        table = table.CopyTo(ws.Cell("D1")).CreateTable();
                        table.Theme = XLTableTheme.TableStyleDark2;

                        table = table.CopyTo(ws.Cell("G1")).CreateTable();
                        table.Theme = XLTableTheme.TableStyleMedium15;

                        wb.SaveAs("tables-create.xlsx");*/

            XLWorkbook workbook = null;

            try
            {
                workbook = new XLWorkbook("tables-create.xlsx");

                var worksheet = workbook.Worksheet(1); // Assuming it's the first worksheet

                // Iterate through the tables in the worksheet
                foreach (var table in worksheet.Tables)
                {
                    // Iterate through the rows and cells of the table
                    foreach (var row in table.DataRange.RowsUsed())
                    {
                        foreach (var cell in row.Cells())
                        {
                            var cellValue = cell.Value.ToString();

                            // Check if cellValue matches the value you're searching for
                            if (cellValue == "5")
                            {
                                // You found a match, do something with it.
                                Console.WriteLine("Found a match in the table: " + cellValue);
                                Thread.Sleep(6000);
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                // Handle exceptions here
            }
            finally
            {
                if (workbook != null)
                {
                    workbook.Dispose();
                }
            }
        }

 

            
/*            public class Pastry
            {
            public string Name { get; set; }
            public int Sales { get; set; }

            public Pastry(string name, int sales)
            {
                Name = name;
                Sales = sales;
            }
        }*/
        }
    }
