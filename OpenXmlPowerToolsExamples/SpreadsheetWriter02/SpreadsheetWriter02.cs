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
                        ColumnHeadings = new[]
                        {
                            new()
                            {
                                Value = "DataType",
                                Bold = true,
                            },
                            new CellDfn
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
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
                                        Value = "Boolean",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
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
                                        Value = "Boolean",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
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
                                        Value = "String",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
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
                                        Value = "int",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 100,
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
                                        Value = "int?",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (int?) 100,
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
                                        Value = "int? (is null)",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
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
                                        Value = "uint",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (uint) 101,
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
                                        Value = "long",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = long.MaxValue,
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
                                        Value = "float",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (float) 123.45,
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
                                        Value = "double",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 123.45,
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
                                        Value = "decimal",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal) 123.45,
                                    },
                                },
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                },
                            },
                        },
                    },
                },
            };

            SpreadsheetWriter.Write(Path.Combine(tempDi.FullName, "Test2.xlsx"), wb);
        }
    }
}
