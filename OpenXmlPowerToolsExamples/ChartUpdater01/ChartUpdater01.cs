using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace OpenXmlPowerTools
{
    class Program
    {
        static void Main(string[] args)
        {
            var n = DateTime.Now;
            var tempDi = new DirectoryInfo(string.Format("ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}", n.Year - 2000, n.Month, n.Day, n.Hour, n.Minute, n.Second));
            tempDi.Create();

            var sourceDi = new DirectoryInfo("../../");
            foreach (var file in sourceDi.GetFiles("*.docx"))
                File.Copy(file.FullName, Path.Combine(tempDi.FullName, file.Name));
            foreach (var file in sourceDi.GetFiles("*.pptx"))
                File.Copy(file.FullName, Path.Combine(tempDi.FullName, file.Name));

            var fileList = Directory.GetFiles(tempDi.FullName, "*.docx");
            foreach (var file in fileList)
            {
                var fi = new FileInfo(file);
                Console.WriteLine(fi.Name);
                var newFileName = "Updated-" + fi.Name;
                var fi2 = new FileInfo(Path.Combine(tempDi.FullName, newFileName));
                File.Copy(fi.FullName, fi2.FullName);

                using (var wDoc = WordprocessingDocument.Open(fi2.FullName, true))
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

            fileList = Directory.GetFiles(tempDi.FullName, "*.pptx");
            foreach (var file in fileList)
            {
                var fi = new FileInfo(file);
                Console.WriteLine(fi.Name);
                var newFileName = "Updated-" + fi.Name;
                var fi2 = new FileInfo(Path.Combine(tempDi.FullName, newFileName));
                File.Copy(fi.FullName, fi2.FullName);

                using (var pDoc = PresentationDocument.Open(fi2.FullName, true))
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
