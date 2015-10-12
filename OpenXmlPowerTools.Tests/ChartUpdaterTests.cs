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
using OpenXmlPowerTools;
using Xunit;

namespace OxPt
{
    public class CuTests
    {
        [Theory]
        [InlineData("CU001-Chart-Cached-Data-01.docx")]
        [InlineData("CU002-Chart-Cached-Data-02.docx")]
        [InlineData("CU003-Chart-Cached-Data-03.docx")]
        [InlineData("CU004-Chart-Cached-Data-04.docx")]
        [InlineData("CU005-Chart-Cached-Data-05.docx")]
        [InlineData("CU006-Chart-Cached-Data-06.docx")]
        [InlineData("CU007-Chart-Cached-Data-07.docx")]
        [InlineData("CU008-Chart-Cached-Data-08.docx")]

        [InlineData("CU009-Chart-Embedded-Xlsx-01.docx")]
        [InlineData("CU010-Chart-Embedded-Xlsx-02.docx")]
        [InlineData("CU011-Chart-Embedded-Xlsx-03.docx")]
        [InlineData("CU012-Chart-Embedded-Xlsx-04.docx")]
        [InlineData("CU013-Chart-Embedded-Xlsx-05.docx")]
        [InlineData("CU014-Chart-Embedded-Xlsx-06.docx")]
        [InlineData("CU015-Chart-Embedded-Xlsx-07.docx")]
        [InlineData("CU016-Chart-Embedded-Xlsx-08.docx")]
        [InlineData("CU017-Chart-Embedded-Xlsx-10.docx")]

        [InlineData("CU018-Chart-Cached-Data-41.pptx")]
        [InlineData("CU019-Chart-Embedded-Xlsx-41.pptx")]

        public void CU001(string name)
        {
            FileInfo templateFile = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            if (templateFile.Extension.ToLower() == ".docx")
            {
                WmlDocument wmlTemplate = new WmlDocument(templateFile.FullName);

                var afterUpdatingDocx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateFile.Name.Replace(".docx", "-processed-by-ChartUpdater.docx")));
                wmlTemplate.SaveAs(afterUpdatingDocx.FullName);

                using (var wDoc = WordprocessingDocument.Open(afterUpdatingDocx.FullName, true))
                {
                    var chart1Data = new ChartData
                    {
                        SeriesNames = new[] {
                            "Car",
                            "Truck",
                            "Van",
                            "Bike",
                            "Boat",
                        },
                        CategoryDataType = ChartDataType.String,
                        CategoryNames = new[] {
                            "Q1",
                            "Q2",
                            "Q3",
                            "Q4",
                        },
                        Values = new double[][] {
                            new double[] {
                                100, 310, 220, 450,
                            },
                            new double[] {
                                200, 300, 350, 411,
                            },
                            new double[] {
                                80, 120, 140, 600,
                            },
                            new double[] {
                                120, 100, 140, 400,
                            },
                            new double[] {
                                200, 210, 210, 480,
                            },
                        },
                    };
                    ChartUpdater.UpdateChart(wDoc, "Chart1", chart1Data);

                    var chart2Data = new ChartData
                    {
                        SeriesNames = new[] {
                            "Series"
                        },
                        CategoryDataType = ChartDataType.String,
                        CategoryNames = new[] {
                                "Cars",
                                "Trucks",
                                "Vans",
                                "Boats",
                            },
                        Values = new double[][] {
                            new double[] {
                                320, 112, 64, 80,
                            },
                        },
                    };
                    ChartUpdater.UpdateChart(wDoc, "Chart2", chart2Data);

                    var chart3Data = new ChartData
                    {
                        SeriesNames = new[] {
                            "X1",
                            "X2",
                            "X3",
                            "X4",
                            "X5",
                            "X6",
                        },
                        CategoryDataType = ChartDataType.String,
                        CategoryNames = new[] {
                            "Y1",
                            "Y2",
                            "Y3",
                            "Y4",
                            "Y5",
                            "Y6",
                        },
                        Values = new double[][] {
                            new double[] {      3.0,      2.1,       .7,      .7,      2.1,      3.0,      },
                            new double[] {      3.0,      2.1,       .8,      .8,      2.1,      3.0,      },
                            new double[] {      3.0,      2.4,      1.2,     1.2,      2.4,      3.0,      },
                            new double[] {      3.0,      2.7,      1.7,     1.7,      2.7,      3.0,      },
                            new double[] {      3.0,      2.9,      2.5,     2.5,      2.9,      3.0,      },
                            new double[] {      3.0,      3.0,      3.0,     3.0,      3.0,      3.0,      },
                        },
                    };
                    ChartUpdater.UpdateChart(wDoc, "Chart3", chart3Data);

                    var chart4Data = new ChartData
                    {
                        SeriesNames = new[] {
                            "Car",
                            "Truck",
                            "Van",
                        },
                        CategoryDataType = ChartDataType.DateTime,
                        CategoryFormatCode = 14,
                        CategoryNames = new[] {
                            ToExcelInteger(new DateTime(2013, 9, 1)),
                            ToExcelInteger(new DateTime(2013, 9, 2)),
                            ToExcelInteger(new DateTime(2013, 9, 3)),
                            ToExcelInteger(new DateTime(2013, 9, 4)),
                            ToExcelInteger(new DateTime(2013, 9, 5)),
                            ToExcelInteger(new DateTime(2013, 9, 6)),
                            ToExcelInteger(new DateTime(2013, 9, 7)),
                            ToExcelInteger(new DateTime(2013, 9, 8)),
                            ToExcelInteger(new DateTime(2013, 9, 9)),
                            ToExcelInteger(new DateTime(2013, 9, 10)),
                            ToExcelInteger(new DateTime(2013, 9, 11)),
                            ToExcelInteger(new DateTime(2013, 9, 12)),
                            ToExcelInteger(new DateTime(2013, 9, 13)),
                            ToExcelInteger(new DateTime(2013, 9, 14)),
                            ToExcelInteger(new DateTime(2013, 9, 15)),
                            ToExcelInteger(new DateTime(2013, 9, 16)),
                            ToExcelInteger(new DateTime(2013, 9, 17)),
                            ToExcelInteger(new DateTime(2013, 9, 18)),
                            ToExcelInteger(new DateTime(2013, 9, 19)),
                            ToExcelInteger(new DateTime(2013, 9, 20)),
                        },
                        Values = new double[][] {
                            new double[] {
                                1, 2, 3, 2, 3, 4, 5, 4, 5, 6, 5, 4, 5, 6, 7, 8, 7, 8, 8, 9,
                            },
                            new double[] {
                                2, 3, 3, 4, 4, 5, 6, 7, 8, 7, 8, 9, 9, 9, 7, 8, 9, 9, 10, 11,
                            },
                            new double[] {
                                2, 3, 3, 3, 3, 2, 2, 2, 3, 2, 3, 3, 4, 4, 4, 3, 4, 5, 5, 4,
                            },
                        },
                    };
                    ChartUpdater.UpdateChart(wDoc, "Chart4", chart4Data);
                }
            }
            if (templateFile.Extension.ToLower() == ".pptx")
            {
                PmlDocument pmlTemplate = new PmlDocument(templateFile.FullName);

                var afterUpdatingPptx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, templateFile.Name.Replace(".pptx", "-processed-by-ChartUpdater.pptx")));
                pmlTemplate.SaveAs(afterUpdatingPptx.FullName);

                using (var pDoc = PresentationDocument.Open(afterUpdatingPptx.FullName, true))
                {
                    var chart1Data = new ChartData
                    {
                        SeriesNames = new[] {
                            "Car",
                            "Truck",
                            "Van",
                        },
                        CategoryDataType = ChartDataType.String,
                        CategoryNames = new[] {
                            "Q1",
                            "Q2",
                            "Q3",
                            "Q4",
                        },
                        Values = new double[][] {
                            new double[] {
                                320, 310, 320, 330,
                            },
                            new double[] {
                                201, 224, 230, 221,
                            },
                            new double[] {
                                180, 200, 220, 230,
                            },
                        },
                    };
                    ChartUpdater.UpdateChart(pDoc, 1, chart1Data);
                }
            }
        }

        private static string ToExcelInteger(DateTime dateTime)
        {
            return dateTime.ToOADate().ToString();
        }
    }
}
