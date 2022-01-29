using Codeuctivity.OpenXmlPowerTools;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xunit;

namespace Codeuctivity.Tests
{
    public class SwTests
    {
        [Fact]
        public void SW001_Simple()
        {
            var wb = new WorkbookDfn
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
            var outXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "SW001-Simple.xlsx"));
            SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        // Breaks with DocumentFormat.OpenXml 2.12  but works till 2.11.3
        [Fact]
        public void SW002_AllDataTypes()
        {
            var wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings = new CellDfn[]
                        {
                            new CellDfn
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
                        Rows = new RowDfn[]
                        {
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "String",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 100,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int?",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 100,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "uint",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 101,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "long",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = long.MaxValue,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "float",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "double",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:str)",
                                    },
                                    new CellDfn {
                                        Value = new DateTime(2012, 1, 8).ToOADate(),
                                        FormatCode= "mm-dd-yy",
                                        Bold = true,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:str)",
                                    },
                                    new CellDfn {
                                        Value = new DateTime(2012, 1, 9).ToOADate(),
                                        FormatCode= "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new CellDfn[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "date (t:d)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 11).ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff"),
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, "SW002-DataTypes.xlsx"));
            SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        private void Validate(FileInfo fi)
        {
            using var sDoc = SpreadsheetDocument.Open(fi.FullName, true);
            var v = new OpenXmlValidator();
            var errors = v.Validate(sDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

            // if a test fails validation post-processing, then can use this code to determine the SDK validation error(s).

            if (errors.Any())
            {
                var sb = new StringBuilder();
                foreach (var item in errors)
                {
                    sb.Append(item.Description).Append(Environment.NewLine);
                }
                var s = sb.ToString();
                Console.WriteLine(s);
            }

            Assert.Empty(errors);
        }

        private static readonly List<string> s_ExpectedErrors = new List<string>()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}