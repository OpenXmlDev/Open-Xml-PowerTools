// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using JetBrains.Annotations;

namespace OpenXmlPowerTools
{
    [PublicAPI]
    public static class ChartUpdater
    {
        private static readonly Dictionary<int, string> FormatCodes = new()
        {
            { 0, "general" },
            { 1, "0" },
            { 2, "0.00" },
            { 3, "#,##0" },
            { 4, "#,##0.00" },
            { 9, "0%" },
            { 10, "0.00%" },
            { 11, "0.00E+00" },
            { 12, "# ?/?" },
            { 13, "# ??/??" },
            { 14, "mm-dd-yy" },
            { 15, "d-mmm-yy" },
            { 16, "d-mmm" },
            { 17, "mmm-yy" },
            { 18, "h:mm AM/PM" },
            { 19, "h:mm:ss AM/PM" },
            { 20, "h:mm" },
            { 21, "h:mm:ss" },
            { 22, "m/d/yy h:mm" },
            { 37, "#,##0 ;(#,##0)" },
            { 38, "#,##0 ;[Red](#,##0)" },
            { 39, "#,##0.00;(#,##0.00)" },
            { 40, "#,##0.00;[Red](#,##0.00)" },
            { 45, "mm:ss" },
            { 46, "[h]:mm:ss" },
            { 47, "mmss.0" },
            { 48, "##0.0E+0" },
            { 49, "@" },
        };

        public static bool UpdateChart(WordprocessingDocument wDoc, string contentControlTag, ChartData chartData)
        {
            MainDocumentPart mainDocumentPart = wDoc.MainDocumentPart ??
                throw new ArgumentException("WordprocessingDocument does not have MainDocumentPart", nameof(wDoc));

            XDocument mdXDoc = mainDocumentPart.GetXDocument();

            XElement cc = mdXDoc
                .Descendants(W.sdt)
                .FirstOrDefault(sdt =>
                    sdt.Elements(W.sdtPr).Elements(W.tag).Attributes(W.val).Any(val => val.Value == contentControlTag));

            if (cc == null)
            {
                return false;
            }

            var chartRid = (string) cc.Descendants(C.chart).Attributes(R.id).FirstOrDefault();

            if (chartRid == null)
            {
                return false;
            }

            var chartPart = (ChartPart) mainDocumentPart.GetPartById(chartRid);
            UpdateChart(chartPart, chartData);

            IEnumerable<XElement> newContent = cc.Elements(W.sdtContent).Elements().Select(e => new XElement(e));
            cc.ReplaceWith(newContent);
            mainDocumentPart.SaveXDocument();

            return true;
        }

        private static void UpdateChart(ChartPart chartPart, ChartData chartData)
        {
            if (chartData.Values.Length != chartData.SeriesNames.Length)
            {
                throw new ArgumentException("Invalid chart data");
            }

            foreach (double[] ser in chartData.Values)
            {
                if (ser.Length != chartData.CategoryNames.Length)
                {
                    throw new ArgumentException("Invalid chart data");
                }
            }

            UpdateSeries(chartPart, chartData);
        }

        private static void UpdateSeries(ChartPart chartPart, ChartData chartData)
        {
            UpdateEmbeddedWorkbook(chartPart, chartData);

            XDocument cpXDoc = chartPart.GetXDocument();
            XElement firstSeries = cpXDoc.Root!.Descendants(C.ser).First();

            string sheetName = null;
            var f = (string) firstSeries.Descendants(C.f).FirstOrDefault();

            if (f != null)
            {
                sheetName = f.Split('!')[0];
            }

            // remove all but first series
            XName chartType = firstSeries.Parent!.Name;
            firstSeries.Parent.Elements(C.ser).Skip(1).Remove();

            IEnumerable<XElement> newSetOfSeries = chartData.SeriesNames
                .Select((_, si) =>
                {
                    XElement cat;
                    XElement oldCat = firstSeries.Elements(C.cat).FirstOrDefault();

                    if (oldCat == null)
                    {
                        throw new OpenXmlPowerToolsException("Invalid chart markup");
                    }

                    bool catHasFormula = oldCat.Descendants(C.f).Any();

                    if (catHasFormula)
                    {
                        XElement newFormula = null;

                        if (sheetName != null)
                        {
                            newFormula = new XElement(C.f, $"{sheetName}!$A$2:$A${chartData.CategoryNames.Length + 1}");
                        }

                        if (chartData.CategoryDataType == ChartDataType.String)
                        {
                            cat = new XElement(C.cat,
                                new XElement(C.strRef,
                                    newFormula,
                                    new XElement(C.strCache,
                                        new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                        chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                            new XAttribute("idx", ci),
                                            new XElement(C.v, chartData.CategoryNames[ci]))))));
                        }
                        else
                        {
                            cat = new XElement(C.cat,
                                new XElement(C.numRef,
                                    newFormula,
                                    new XElement(C.numCache,
                                        new XElement(C.formatCode, FormatCodes[chartData.CategoryFormatCode]),
                                        new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                        chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                            new XAttribute("idx", ci),
                                            new XElement(C.v, chartData.CategoryNames[ci]))))));
                        }
                    }
                    else
                    {
                        if (chartData.CategoryDataType == ChartDataType.String)
                        {
                            cat = new XElement(C.cat,
                                new XElement(C.strLit,
                                    new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                    chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                        new XAttribute("idx", ci),
                                        new XElement(C.v, chartData.CategoryNames[ci])))));
                        }
                        else
                        {
                            cat = new XElement(C.cat,
                                new XElement(C.numLit,
                                    new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                    chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                        new XAttribute("idx", ci),
                                        new XElement(C.v, chartData.CategoryNames[ci])))));
                        }
                    }

                    XElement newCval;

                    if (sheetName == null)
                    {
                        newCval = new XElement(C.val,
                            new XElement(C.numLit,
                                new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                    new XAttribute("idx", ci),
                                    new XElement(C.v, chartData.Values[si][ci])))));
                    }
                    else
                    {
                        XElement numRef = firstSeries.Elements(C.val).Elements(C.numRef).First();

                        newCval = new XElement(C.val,
                            new XElement(C.numRef,
                                new XElement(C.f,
                                    string.Format("{0}!${2}$2:${2}${1}", sheetName, chartData.CategoryNames.Length + 1,
                                        SpreadsheetMLUtil.IntToColumnId(si + 1))),
                                new XElement(C.numCache,
                                    numRef.Descendants(C.formatCode),
                                    new XElement(C.ptCount, new XAttribute("val", chartData.CategoryNames.Length)),
                                    chartData.CategoryNames.Select((_, ci) => new XElement(C.pt,
                                        new XAttribute("idx", ci),
                                        new XElement(C.v, chartData.Values[si][ci]))))));
                    }

                    XElement tx;
                    bool serHasFormula = firstSeries.Descendants(C.f).Any();

                    if (serHasFormula)
                    {
                        XElement newFormula = null;

                        if (sheetName != null)
                        {
                            newFormula = new XElement(C.f, $"{sheetName}!${SpreadsheetMLUtil.IntToColumnId(si + 1)}$1");
                        }

                        tx = new XElement(C.tx,
                            new XElement(C.strRef,
                                newFormula,
                                new XElement(C.strCache,
                                    new XElement(C.ptCount, new XAttribute("val", 1)),
                                    new XElement(C.pt,
                                        new XAttribute("idx", 0),
                                        new XElement(C.v, chartData.SeriesNames[si])))));
                    }
                    else
                    {
                        tx = new XElement(C.tx,
                            new XElement(C.v, chartData.SeriesNames[si]));
                    }

                    XElement newSer = null;

                    if (chartType == C.area3DChart || chartType == C.areaChart)
                    {
                        newSer = new XElement(C.ser,
                            // common
                            new XElement(C.idx, new XAttribute("val", si)),
                            new XElement(C.order, new XAttribute("val", si)),
                            tx,
                            firstSeries.Elements(C.spPr),

                            // CT_AreaSer
                            firstSeries.Elements(C.pictureOptions),
                            firstSeries.Elements(C.dPt),
                            firstSeries.Elements(C.dLbls),
                            firstSeries.Elements(C.trendline),
                            firstSeries.Elements(C.errBars),
                            cat,
                            newCval,
                            firstSeries.Elements(C.extLst));
                    }
                    else if (chartType == C.bar3DChart || chartType == C.barChart)
                    {
                        newSer = new XElement(C.ser,
                            // common
                            new XElement(C.idx, new XAttribute("val", si)),
                            new XElement(C.order, new XAttribute("val", si)),
                            tx,
                            firstSeries.Elements(C.spPr),

                            // CT_BarSer
                            firstSeries.Elements(C.invertIfNegative),
                            firstSeries.Elements(C.pictureOptions),
                            firstSeries.Elements(C.dPt),
                            firstSeries.Elements(C.dLbls),
                            firstSeries.Elements(C.trendline),
                            firstSeries.Elements(C.errBars),
                            cat,
                            newCval,
                            firstSeries.Elements(C.shape),
                            firstSeries.Elements(C.extLst));
                    }
                    else if (chartType == C.line3DChart || chartType == C.lineChart || chartType == C.stockChart)
                    {
                        newSer = new XElement(C.ser,
                            // common
                            new XElement(C.idx, new XAttribute("val", si)),
                            new XElement(C.order, new XAttribute("val", si)),
                            tx,
                            firstSeries.Elements(C.spPr),

                            // CT_LineSer
                            firstSeries.Elements(C.marker),
                            firstSeries.Elements(C.dPt),
                            firstSeries.Elements(C.dLbls),
                            firstSeries.Elements(C.trendline),
                            firstSeries.Elements(C.errBars),
                            cat,
                            newCval,
                            firstSeries.Elements(C.smooth),
                            firstSeries.Elements(C.extLst));
                    }
                    else if (chartType == C.doughnutChart ||
                             chartType == C.ofPieChart ||
                             chartType == C.pie3DChart ||
                             chartType == C.pieChart)
                    {
                        newSer = new XElement(C.ser,
                            // common
                            new XElement(C.idx, new XAttribute("val", si)),
                            new XElement(C.order, new XAttribute("val", si)),
                            tx,
                            firstSeries.Elements(C.spPr),

                            // CT_PieSer
                            firstSeries.Elements(C.explosion),
                            firstSeries.Elements(C.dPt),
                            firstSeries.Elements(C.dLbls),
                            cat,
                            newCval,
                            firstSeries.Elements(C.extLst));
                    }
                    else if (chartType == C.surface3DChart || chartType == C.surfaceChart)
                    {
                        newSer = new XElement(C.ser,
                            // common
                            new XElement(C.idx, new XAttribute("val", si)),
                            new XElement(C.order, new XAttribute("val", si)),
                            tx,
                            firstSeries.Elements(C.spPr),

                            // CT_SurfaceSer
                            cat,
                            newCval,
                            firstSeries.Elements(C.extLst));
                    }

                    if (newSer == null)
                    {
                        throw new OpenXmlPowerToolsException("Unsupported chart type");
                    }

                    int accentNumber = si % 6 + 1;
                    newSer = (XElement) UpdateAccentTransform(newSer, accentNumber);
                    return newSer;
                });

            firstSeries.ReplaceWith(newSetOfSeries);
            chartPart.SaveXDocument();
        }

        private static void UpdateEmbeddedWorkbook(ChartPart chartPart, ChartData chartData)
        {
            XDocument cpXDoc = chartPart.GetXDocument();
            XElement root = cpXDoc.Root!;
            XElement firstSeries = root.Descendants(C.ser).FirstOrDefault();

            if (firstSeries == null)
            {
                return;
            }

            var firstFormula = (string) firstSeries.Descendants(C.f).FirstOrDefault();

            if (firstFormula == null)
            {
                return;
            }

            string sheet = firstFormula.Split('!')[0];

            var embeddedSpreadsheetRid = (string) root.Descendants(C.externalData).Attributes(R.id).FirstOrDefault();

            if (embeddedSpreadsheetRid == null)
            {
                return;
            }

            OpenXmlPart embeddedSpreadsheet = chartPart.GetPartById(embeddedSpreadsheetRid);

            using SpreadsheetDocument sDoc = SpreadsheetDocument.Open(embeddedSpreadsheet.GetStream(), true);

            WorkbookPart workbookPart = sDoc.WorkbookPart!;
            XElement wbRoot = workbookPart.GetXElement()!;

            var sheetRid = (string) wbRoot
                .Elements(S.sheets)
                .Elements(S.sheet)
                .Where(s => (string) s.Attribute("name") == sheet)
                .Attributes(R.id)
                .FirstOrDefault();

            if (sheetRid == null)
            {
                return;
            }

            OpenXmlPart sheetPart = workbookPart.GetPartById(sheetRid);
            XDocument xdSheet = sheetPart.GetXDocument();
            XElement sheetData = xdSheet.Descendants(S.sheetData).First();

            WorkbookStylesPart stylePart = workbookPart.WorkbookStylesPart!;
            XDocument xdStyle = stylePart.GetXDocument();

            var categoryStyleId = 0;

            if (chartData.CategoryFormatCode != 0)
            {
                categoryStyleId = AddDxfToDxfs(xdSheet, xdStyle, chartData.CategoryFormatCode);
            }

            stylePart.SaveXDocument();

            var firstRow = new XElement(S.row,
                new XAttribute("r", "1"),
                new XAttribute("spans", $"1:{chartData.SeriesNames.Length + 1}"),
                new[]
                {
                    new XElement(S.c,
                        new XAttribute("r", "A1"),
                        new XAttribute("t", "str"),
                        new XElement(S.v,
                            new XAttribute(XNamespace.Xml + "space", "preserve"),
                            " ")),
                }.Concat(chartData.SeriesNames
                    .Select((sn, i) => new XElement(S.c,
                        new XAttribute("r", RowColToString(0, i + 1)),
                        new XAttribute("t", "str"),
                        new XElement(S.v, sn)))));

            IEnumerable<XElement> otherRows = chartData
                .CategoryNames
                .Select((cn, r) =>
                {
                    var row = new XElement(S.row,
                        new XAttribute("r", r + 2),
                        new XAttribute("spans", $"1:{chartData.SeriesNames.Length + 1}"),
                        new[]
                        {
                            new XElement(S.c,
                                new XAttribute("r", RowColToString(r + 1, 0)),
                                categoryStyleId != 0 ? new XAttribute("s", categoryStyleId) : null,
                                chartData.CategoryDataType == ChartDataType.String
                                    ? new XAttribute("t", "str")
                                    : null,
                                new XElement(S.v, cn)),
                        }.Concat(Enumerable.Range(0, chartData.Values.Length)
                            .Select((_, ci) =>
                            {
                                var cell = new XElement(S.c,
                                    new XAttribute("r", RowColToString(r + 1, ci + 1)),
                                    new XElement(S.v, chartData.Values[ci][r]));

                                return cell;
                            })));

                    return row;
                });

            IEnumerable<XElement> allRows = new[] { firstRow }.Concat(otherRows);

            var newSheetData = new XElement(S.sheetData, allRows);
            sheetData.ReplaceWith(newSheetData);
            sheetPart.SaveXDocument();

            var tablePartRid = (string) xdSheet
                .Root!
                .Elements(S.tableParts)
                .Elements(S.tablePart)
                .Attributes(R.id)
                .FirstOrDefault();

            if (tablePartRid != null)
            {
                OpenXmlPart partTable = sheetPart.GetPartById(tablePartRid);
                XDocument xdTablePart = partTable.GetXDocument();
                XAttribute xaRef = xdTablePart.Root!.Attribute("ref")!;

                xaRef.Value =
                    $"A1:{RowColToString(chartData.CategoryNames.Length - 1, chartData.SeriesNames.Length)}";

                var xeNewTableColumns = new XElement(S.tableColumns,
                    new XAttribute("count", chartData.SeriesNames.Length + 1),
                    new[]
                    {
                        new XElement(S.tableColumn,
                            new XAttribute("id", 1),
                            new XAttribute("name", " ")),
                    }.Concat(chartData.SeriesNames.Select((cn, ci) =>
                        new XElement(S.tableColumn,
                            new XAttribute("id", ci + 2),
                            new XAttribute("name", cn)))));

                XElement xeExistingTableColumns = xdTablePart.Root.Element(S.tableColumns);

                if (xeExistingTableColumns != null)
                {
                    xeExistingTableColumns.ReplaceWith(xeNewTableColumns);
                }

                partTable.SaveXDocument();
            }
        }

        private static int AddDxfToDxfs(XDocument xdSheet, XDocument xdStyle, int formatCodeToAdd)
        {
            // add xf to cellXfs
            XElement cellXfs = xdStyle.Root!.Element(S.cellXfs);

            if (cellXfs == null)
            {
                XElement cellStyleXfs = xdStyle.Root.Element(S.cellStyleXfs);

                if (cellStyleXfs != null)
                {
                    cellStyleXfs.AddAfterSelf(new XElement(S.cellXfs, new XAttribute("count", 0)));
                    cellXfs = xdSheet.Root!.Element(S.cellXfs);
                }
            }

            if (cellXfs == null)
            {
                XElement borders = xdStyle.Root.Element(S.borders);

                if (borders != null)
                {
                    borders.AddAfterSelf(new XElement(S.cellXfs, new XAttribute("count", 0)));
                    cellXfs = xdSheet.Root!.Element(S.cellXfs);
                }
            }

            if (cellXfs == null)
            {
                throw new OpenXmlPowerToolsException("Internal error");
            }

            var cnt = (int) cellXfs.Attribute("count");
            cnt++;

            // cellXfs.Attribute("count").Value = cnt.ToString();
            cellXfs.SetAttributeValue("count", cnt);

            cellXfs.Add(new XElement(S.xf,
                new XAttribute("numFmtId", formatCodeToAdd),
                new XAttribute("fontId", 0),
                new XAttribute("fillId", 0),
                new XAttribute("borderId", 0),
                new XAttribute("applyNumberFormat", 1)));

            return cnt - 1;
        }

        private static string RowColToString(int row, int col)
        {
            return SpreadsheetMLUtil.IntToColumnId(col) + (row + 1);
        }

        private static object UpdateAccentTransform(XNode node, int accentNumber)
        {
            return node is not XElement element
                ? node
                : element.Name == A.schemeClr && (string) element.Attribute("val") == "accent1"
                    ? new XElement(A.schemeClr, new XAttribute("val", "accent" + accentNumber))
                    : new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => UpdateAccentTransform(n, accentNumber)));
        }

        public static bool UpdateChart(PresentationDocument pDoc, int slideNumber, ChartData chartData)
        {
            PresentationPart presentationPart = pDoc.PresentationPart;

            if (presentationPart is null)
            {
                throw new ArgumentException("Presentation does not have PresentationPart", nameof(pDoc));
            }

            XElement sldIdElement = presentationPart.GetXDocument()
                .Elements(P.presentation)
                .Elements(P.sldIdLst)
                .Elements(P.sldId)
                .Skip(slideNumber - 1)
                .FirstOrDefault();

            if (sldIdElement == null)
            {
                return false;
            }

            var slideRId = (string) sldIdElement.Attribute(R.id);
            OpenXmlPart slidePart = presentationPart.GetPartById(slideRId);

            var chartRid = (string) slidePart.GetXDocument()
                .Descendants(C.chart)
                .Attributes(R.id)
                .FirstOrDefault();

            if (chartRid == null)
            {
                // TODO: Revisit. This returned true, which did not seem to make sense.
                return false;
            }

            var chartPart = (ChartPart) slidePart.GetPartById(chartRid);
            UpdateChart(chartPart, chartData);
            return true;
        }
    }
}
