using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace SpreadsheetWriterExample
{
    class Program
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            WorkbookDfn wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
                            {
                                Value = "Name",
                                Bold = true,
                            },
                            new CellDfn
                            {
                                Value = "Age",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            },
                            new CellDfn
                            {
                                Value = "Rate",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Left,
                            }
                        },
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 50,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)45.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 42,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)78.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            SpreadsheetWriter.Write(Path.Combine(tempDi.FullName, "Test1.xlsx"), wb);
        }
    }
}
