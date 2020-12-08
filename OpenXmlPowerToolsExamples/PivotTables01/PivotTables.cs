using OpenXmlPowerTools;
using System;
using System.IO;

namespace ExamplePivotTables
{
    internal class PivotTableExample
    {
        private static void Main()
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            // Update an existing pivot table
            var qs = new FileInfo("../../QuarterlySales.xlsx");
            var qsu = new FileInfo(Path.Combine(tempDi.FullName, "QuarterlyPivot.xlsx"));

            var row = 1;
            using (var streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName(qs.FullName)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var sheet = WorksheetAccessor.GetWorksheet(doc, "Range");
                    using (var source = new StreamReader("../../PivotData.txt"))
                    {
                        while (!source.EndOfStream)
                        {
                            var line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                var fields = line.Split(',');
                                var column = 1;
                                foreach (var item in fields)
                                {
                                    if (double.TryParse(item, out var num))
                                    {
                                        WorksheetAccessor.SetCellValue(sheet, row, column++, num);
                                    }
                                    else
                                    {
                                        WorksheetAccessor.SetCellValue(sheet, row, column++, item);
                                    }
                                }
                            }
                            row++;
                        }
                    }
                    sheet.PutXDocument();

                    WorksheetAccessor.UpdateRangeEndRow(doc, "Sales", row - 1);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(qsu.FullName);
            }

            // Create from scratch
            row = 1;
            var maxColumn = 1;
            using (var streamDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument())
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.CreateDefaultStyles(doc);
                    var sheet = WorksheetAccessor.AddWorksheet(doc, "Range");
                    var ms = new MemorySpreadsheet();

                    var southIndex = WorksheetAccessor.GetStyleIndex(doc, 0, 8, 1, 2,
                        new WorksheetAccessor.CellAlignment { HorizontalAlignment = WorksheetAccessor.CellAlignment.Horizontal.Center },
                        true, false);
                    var gradient = new WorksheetAccessor.GradientFill(90);
                    gradient.AddStop(new WorksheetAccessor.GradientStop(0, new WorksheetAccessor.ColorInfo("FF92D050")));
                    gradient.AddStop(new WorksheetAccessor.GradientStop(1, new WorksheetAccessor.ColorInfo("FF0070C0")));
                    var northIndex = WorksheetAccessor.GetStyleIndex(doc, 0,
                        WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                        {
                            Italic = true,
                            Size = 8,
                            Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 1),
                            Name = "Times New Roman",
                            Family = 1
                        }),
                    WorksheetAccessor.GetFillIndex(doc, gradient),
                    WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        DiagonalDown = true,
                        Diagonal =
                            new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF616100"))
                    }),
                    null, false, false);
                    WorksheetAccessor.CheckNumberFormat(doc, 100, "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");
                    var amountIndex = WorksheetAccessor.GetStyleIndex(doc, 100, 0, 0, 0, null, false, false);

                    using (var source = new StreamReader("../../PivotData.txt"))
                    {
                        while (!source.EndOfStream)
                        {
                            var line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                var fields = line.Split(',');
                                var column = 1;
                                foreach (var item in fields)
                                {
                                    if (double.TryParse(item, out var num))
                                    {
                                        if (column == 6)
                                        {
                                            ms.SetCellValue(row, column++, num, amountIndex);
                                        }
                                        else
                                        {
                                            ms.SetCellValue(row, column++, num);
                                        }
                                    }
                                    else if (item == "Accessories")
                                    {
                                        ms.SetCellValue(row, column++, item, WorksheetAccessor.GetStyleIndex(doc, "Good"));
                                    }
                                    else if (item == "South")
                                    {
                                        ms.SetCellValue(row, column++, item, southIndex);
                                    }
                                    else if (item == "North")
                                    {
                                        ms.SetCellValue(row, column++, item, northIndex);
                                    }
                                    else
                                    {
                                        ms.SetCellValue(row, column++, item);
                                    }
                                }
                                maxColumn = column - 1;
                            }
                            row++;
                        }
                    }
                    WorksheetAccessor.SetSheetContents(sheet, ms);
                    WorksheetAccessor.SetRange(doc, "Sales", "Range", 1, 1, row - 1, maxColumn);
                    var pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);

                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Amount");
                    WorksheetAccessor.AddPivotAxis(pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "NewPivot.xlsx"));
            }

            // Add pivot table to existing spreadsheet
            // Demonstrate multiple data fields
            using (var streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName("../../QuarterlyUnitSales.xlsx")))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);

                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Total");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Quantity");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Unit Price");
                    WorksheetAccessor.AddPivotAxis(pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "QuarterlyUnitSalesWithPivot.xlsx"));
            }
        }
    }
}