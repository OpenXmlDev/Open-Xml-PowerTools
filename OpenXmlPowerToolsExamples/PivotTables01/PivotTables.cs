using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Timers;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace ExamplePivotTables
{
    class PivotTableExample
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            // Update an existing pivot table
            FileInfo qs = new FileInfo("../../QuarterlySales.xlsx");
            FileInfo qsu = new FileInfo(Path.Combine(tempDi.FullName, "QuarterlyPivot.xlsx"));

            int row = 1;
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName(qs.FullName)))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetPart sheet = WorksheetAccessor.GetWorksheet(doc, "Range");
                    using (StreamReader source = new StreamReader("../../PivotData.txt"))
                    {
                        while (!source.EndOfStream)
                        {
                            string line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                string[] fields = line.Split(',');
                                int column = 1;
                                foreach (string item in fields)
                                {
                                    double num;
                                    if (double.TryParse(item, out num))
                                        WorksheetAccessor.SetCellValue(doc, sheet, row, column++, num);
                                    else
                                        WorksheetAccessor.SetCellValue(doc, sheet, row, column++, item);
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
            int maxColumn = 1;
            using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument())
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.CreateDefaultStyles(doc);
                    WorksheetPart sheet = WorksheetAccessor.AddWorksheet(doc, "Range");
                    MemorySpreadsheet ms = new MemorySpreadsheet();

#if false
                    int font0 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 1),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font2 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 18,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 3),
                        Name = "Cambria",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Major
                    });
                    int font3 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 15,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 3),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font4 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 13,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 3),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font5 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 3),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font6 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF006100"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font7 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF9C0006"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font8 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF9C6500"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font9 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF3F3F76"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font10 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF3F3F3F"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font11 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FFFA7D00"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font12 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FFFA7D00"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font13 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 0),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font14 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FFFF0000"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font15 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Italic = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo("FF7F7F7F"),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font16 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Bold = true,
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 1),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });
                    int font17 = WorksheetAccessor.GetFontIndex(doc, new WorksheetAccessor.Font
                    {
                        Size = 11,
                        Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 0),
                        Name = "Calibri",
                        Family = 2,
                        Scheme = WorksheetAccessor.Font.SchemeType.Minor
                    });

                    int fill0 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.None, null, null));
                    int fill1 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Gray125, null, null));
                    int fill2 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFC6EFCE")));
                    int fill3 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFFFC7CE")));
                    int fill4 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFFFEB9C")));
                    int fill5 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFFFCC99")));
                    int fill6 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFF2F2F2")));
                    int fill7 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFA5A5A5")));
                    int fill8 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo("FFFFFFCC")));
                    int fill9 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 4)));
                    int fill10 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(4, 0.79998168889431442)));
                    int fill11 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(4, 0.59999389629810485)));
                    int fill12 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(4, 0.39997558519241921)));
                    int fill13 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 5)));
                    int fill14 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(5, 0.79998168889431442)));
                    int fill15 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(5, 0.59999389629810485)));
                    int fill16 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(5, 0.39997558519241921)));
                    int fill17 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 6)));
                    int fill18 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(6, 0.79998168889431442)));
                    int fill19 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(6, 0.59999389629810485)));
                    int fill20 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(6, 0.39997558519241921)));
                    int fill21 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 7)));
                    int fill22 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(7, 0.79998168889431442)));
                    int fill23 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(7, 0.59999389629810485)));
                    int fill24 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(7, 0.39997558519241921)));
                    int fill25 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 8)));
                    int fill26 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(8, 0.79998168889431442)));
                    int fill27 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(8, 0.59999389629810485)));
                    int fill28 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(8, 0.39997558519241921)));
                    int fill29 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        null, new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 9)));
                    int fill30 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(9, 0.79998168889431442)));
                    int fill31 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(9, 0.59999389629810485)));
                    int fill32 = WorksheetAccessor.GetFillIndex(doc, new WorksheetAccessor.PatternFill(WorksheetAccessor.PatternFill.PatternType.Solid,
                        new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Indexed, 65),
                        new WorksheetAccessor.ColorInfo(9, 0.39997558519241921)));

                    int border1 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thick,
                            new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 4))
                    });
                    int border2 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thick, new WorksheetAccessor.ColorInfo(4, 0.499984740745262))
                    });
                    int border3 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Medium, new WorksheetAccessor.ColorInfo(4, 0.39997558519241921))
                    });
                    int border4 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Left = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF7F7F7F")),
                        Right = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF7F7F7F")),
                        Top = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF7F7F7F")),
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF7F7F7F"))
                    });
                    int border5 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Left = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Right = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Top = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FF3F3F3F"))
                    });
                    int border6 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double, new WorksheetAccessor.ColorInfo("FFFF8001"))
                    });
                    int border7 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Left = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Right = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Top = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double, new WorksheetAccessor.ColorInfo("FF3F3F3F")),
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double, new WorksheetAccessor.ColorInfo("FF3F3F3F"))
                    });
                    int border8 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Left = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FFB2B2B2")),
                        Right = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FFB2B2B2")),
                        Top = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FFB2B2B2")),
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin, new WorksheetAccessor.ColorInfo("FFB2B2B2"))
                    });
                    int border9 = WorksheetAccessor.GetBorderIndex(doc, new WorksheetAccessor.Border
                    {
                        Top = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Thin,
                            new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 4)),
                        Bottom = new WorksheetAccessor.BorderLine(WorksheetAccessor.BorderLine.LineStyle.Double,
                            new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 4))
                    });
#endif

                    int southIndex = WorksheetAccessor.GetStyleIndex(doc, 0, 8, 1, 2,
                        new WorksheetAccessor.CellAlignment { HorizontalAlignment = WorksheetAccessor.CellAlignment.Horizontal.Center },
                        true, false);
                    WorksheetAccessor.GradientFill gradient = new WorksheetAccessor.GradientFill(90);
                    gradient.AddStop(new WorksheetAccessor.GradientStop(0, new WorksheetAccessor.ColorInfo("FF92D050")));
                    gradient.AddStop(new WorksheetAccessor.GradientStop(1, new WorksheetAccessor.ColorInfo("FF0070C0")));
                    int northIndex = WorksheetAccessor.GetStyleIndex(doc, 0,
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
                    int amountIndex = WorksheetAccessor.GetStyleIndex(doc, 100, 0, 0, 0, null, false, false);

                    using (StreamReader source = new StreamReader("../../PivotData.txt"))
                    {
                        while (!source.EndOfStream)
                        {
                            string line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                string[] fields = line.Split(',');
                                int column = 1;
                                foreach (string item in fields)
                                {
                                    double num;
                                    if (double.TryParse(item, out num))
                                    {
                                        if (column == 6)
                                            ms.SetCellValue(row, column++, num, amountIndex);
                                        else
                                            ms.SetCellValue(row, column++, num);
                                    }
                                    else if (item == "Accessories")
                                        ms.SetCellValue(row, column++, item, WorksheetAccessor.GetStyleIndex(doc, "Good"));
                                    else if (item == "South")
                                        ms.SetCellValue(row, column++, item, southIndex);
                                    else if (item == "North")
                                        ms.SetCellValue(row, column++, item, northIndex);
                                    else
                                        ms.SetCellValue(row, column++, item);
                                }
                                maxColumn = column - 1;
                            }
                            row++;
                        }
                    }
                    WorksheetAccessor.SetSheetContents(doc, sheet, ms);
                    WorksheetAccessor.SetRange(doc, "Sales", "Range", 1, 1, row - 1, maxColumn);
                    WorksheetPart pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);

                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Amount");
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "NewPivot.xlsx"));
            }


            // Add pivot table to existing spreadsheet
            // Demonstrate multiple data fields
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(
                SmlDocument.FromFileName("../../QuarterlyUnitSales.xlsx")))
            {
                using (SpreadsheetDocument doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetPart pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);

                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Total");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Quantity");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Unit Price");
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }
                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(tempDi.FullName, "QuarterlyUnitSalesWithPivot.xlsx"));
            }
        }
    }
}
