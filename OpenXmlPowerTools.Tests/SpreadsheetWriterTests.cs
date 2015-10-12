/***************************************************************************

Copyright (c) Microsoft Corporation 2012-2015.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Published at http://OpenXmlDeveloper.org
Resource Center and Documentation: http://openxmldeveloper.org/wiki/w/wiki/powertools-for-open-xml.aspx

Developer: Eric White
Blog: http://www.ericwhite.com
Twitter: @EricWhiteDev
Email: eric@ericwhite.com

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Sw = OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class SwTests
    {
        [Fact]
        public void SW001_Simple()
        {
            Sw.WorkbookDfn wb = new Sw.WorkbookDfn
            {
                Worksheets = new Sw.WorksheetDfn[]
                {
                    new Sw.WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new Sw.CellDfn[]
                        {
                            new Sw.CellDfn
                            {
                                Value = "Name",
                                Bold = true,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Age",
                                Bold = true,
                                HorizontalCellAlignment = Sw.HorizontalCellAlignment.Left,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Rate",
                                Bold = true,
                                HorizontalCellAlignment = Sw.HorizontalCellAlignment.Left,
                            }
                        },
                        Rows = new Sw.RowDfn[]
                        {
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = 50,
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)45.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = 42,
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)78.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(Sw.TestUtil.TempDir.FullName, "SW001-Simple.xlsx"));
            Sw.SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        [Fact]
        public void SW002_AllDataTypes()
        {
            Sw.WorkbookDfn wb = new Sw.WorkbookDfn
            {
                Worksheets = new Sw.WorksheetDfn[]
                {
                    new Sw.WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings = new Sw.CellDfn[]
                        {
                            new Sw.CellDfn
                            {
                                Value = "DataType",
                                Bold = true,
                            },
                            new Sw.CellDfn
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = Sw.HorizontalCellAlignment.Right,
                            },
                        },
                        Rows = new Sw.RowDfn[]
                        {
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "String",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = Sw.HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (int)100,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int?",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "uint",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "long",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = Int64.MaxValue,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "float",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "double",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (double)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new Sw.RowDfn
                            {
                                Cells = new Sw.CellDfn[]
                                {
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new Sw.CellDfn {
                                        CellDataType = Sw.CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = Sw.HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            var outXlsx = new FileInfo(Path.Combine(Sw.TestUtil.TempDir.FullName, "SW002-DataTypes.xlsx"));
            Sw.SpreadsheetWriter.Write(outXlsx.FullName, wb);
            Validate(outXlsx);
        }

        private void Validate(FileInfo fi)
        {
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(fi.FullName, true))
            {
                OpenXmlValidator v = new OpenXmlValidator();
                var errors = v.Validate(sDoc).Where(ve => !s_ExpectedErrors.Contains(ve.Description));

#if false
                // if a test fails validation post-processing, then can use this code to determine the SDK
                // validation error(s).

                if (errors.Count() != 0)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (var item in errors)
                    {
                        sb.Append(item.Description).Append(Environment.NewLine);
                    }
                    var s = sb.ToString();
                    Console.WriteLine(s);
                }
#endif

                Assert.Equal(0, errors.Count());
            }
        }

        private static List<string> s_ExpectedErrors = new List<string>()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}
