// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.IO;
using OpenXmlPowerTools;

namespace SpreadsheetWriterExample
{
    internal static class Program
    {
        private static void Main()
        {
            DateTime n = DateTime.Now;

            var tempDi = new DirectoryInfo(
                $"ExampleOutput-{n.Year - 2000:00}-{n.Month:00}-{n.Day:00}-{n.Hour:00}{n.Minute:00}{n.Second:00}");

            tempDi.Create();

            var wb = new WorkbookDfn
            {
                Worksheets = new[]
                {
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new[]
                        {
                            new()
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
                            },
                        },
                        Rows = new[]
                        {
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 50,
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal) 45.00,
                                        FormatCode = "0.00",
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 42,
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal) 78.00,
                                        FormatCode = "0.00",
                                    },
                                },
                            },
                        },
                    },
                },
            };

            SpreadsheetWriter.Write(Path.Combine(tempDi.FullName, "Test1.xlsx"), wb);
        }
    }
}
