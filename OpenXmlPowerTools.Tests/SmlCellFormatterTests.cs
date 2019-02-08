// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class CfTests
    {
        [Theory]
        [InlineData("General", "0", "0", null)]
        [InlineData("0", "1.1000000000000001", "1", null)]
        [InlineData("0", "10.1", "10", null)]
        [InlineData("0", "100.1", "100", null)]
        [InlineData("0", "100000000.09999999", "100000000", null)]
        [InlineData("0.00", "1.1000000000000001", "1.10", null)]
        [InlineData("0.00", "10.1", "10.10", null)]
        [InlineData("0.00", "100000000.09999999", "100000000.10", null)]
        [InlineData("#,##0", "1.1000000000000001", "1", null)]
        [InlineData("#,##0", "10.1", "10", null)]
        [InlineData("#,##0", "100000000.09999999", "100,000,000", null)]
        [InlineData("#,##0", "1000000000.1", "1,000,000,000", null)]
        [InlineData("#,##0.00", "1.1000000000000001", "1.10", null)]
        [InlineData("#,##0.00", "10.1", "10.10", null)]
        [InlineData("#,##0.00", "1000.1", "1,000.10", null)]
        [InlineData("#,##0.00", "100000000.09999999", "100,000,000.10", null)]
        [InlineData("0%", "0.01", "1%", null)]
        [InlineData("0%", "0.25", "25%", null)]
        [InlineData("0%", "1", "100%", null)]
        [InlineData("0%", "2", "200%", null)]
        [InlineData("0%", "0.1", "10%", null)]
        [InlineData("0.00%", "0.01", "1.00%", null)]
        [InlineData("0.00%", "0.25", "25.00%", null)]
        [InlineData("0.00%", "1", "100.00%", null)]
        [InlineData("0.00%", "2", "200.00%", null)]
        [InlineData("0.00%", "0.1", "10.00%", null)]
        [InlineData("0.00%", "0.1025", "10.25%", null)]
        [InlineData("0.00E+00", "0.01", "1.00E-02", null)]
        [InlineData("0.00E+00", "0.25", "2.50E-01", null)]
        [InlineData("0.00E+00", "1", "1.00E+00", null)]
        [InlineData("0.00E+00", "100", "1.00E+02", null)]
        [InlineData("0.00E+00", "1000", "1.00E+03", null)]
        [InlineData("0.00E+00", "10000.1", "1.00E+04", null)]
        [InlineData("0.00E+00", "100000.5", "1.00E+05", null)]
        [InlineData("0.00E+00", "0.1", "1.00E-01", null)]
        [InlineData("# ?/?", "0.125", "0.13", null)]
        [InlineData("# ?/?", "0.25", "0.25", null)]
        [InlineData("# ??/??", "0.125", "0.13", null)]
        [InlineData("# ??/??", "0.25", "0.25", null)]
        [InlineData("mm-dd-yy", "42344", "12-06-15", null)]
        [InlineData("d-mmm-yy", "42344", "6-Dec-15", null)]
        [InlineData("d-mmm", "42344", "6-Dec", null)]
        [InlineData("mmm-yy", "42344", "Dec-15", null)]
        [InlineData("h:mm AM/PM", "42344.295138888891", "7:05 AM", null)]
        [InlineData("h:mm:ss AM/PM", "42344.295405092591", "7:05:23 AM", null)]
        [InlineData("h:mm", "42344.295405092591", "7:05", null)]
        [InlineData("h:mm:ss", "42344.295405092591", "7:05:23", null)]
        [InlineData("m/d/yy h:mm", "42344.295405092591", "12/6/15 7:05", null)]
        [InlineData("#,##0 ;(#,##0)", "100", "100", null)]
        [InlineData("#,##0 ;(#,##0)", "-100", "(100)", null)]
        [InlineData("#,##0 ;[Red](#,##0)", "100", "100", null)]
        [InlineData("#,##0 ;[Red](#,##0)", "-100", "(100)", "Red")]
        [InlineData("#,##0.00;(#,##0.00)", "100.00", "100.00", null)]
        [InlineData("#,##0.00;(#,##0.00)", "-100.00", "(100.00)", null)]
        [InlineData("#,##0.00;[Red](#,##0.00)", "100.00", "100.00", null)]
        [InlineData("#,##0.00;[Red](#,##0.00)", "-100.00", "(100.00)", "Red")]
        [InlineData("mm:ss", "42344.295405092591", "05:23", null)]
        [InlineData("[h]:mm:ss", "42344.295405092591", "1016263:05:23", null)]
        [InlineData("mm:ss.0", "42344.295445092591", "05:26:456", null)]
        [InlineData("##0.0E+0", "100.0", "100.0E+0", null)]
        [InlineData("##0.0E+0", "543.210", "543.2E+0", null)]
        
        public void CF001(string formatCode, string value, string expected, string expectedColor)
        {
            string color;
            string r = SmlCellFormatter.FormatCell(formatCode, value, out color);
            Assert.Equal(expected, r);
            Assert.Equal(expectedColor, color);
        }

        [Theory]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "A1:A1", "$123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "A2:A2", "-$123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "A3:A3", "$0.00", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "B1:B1", "$ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "B2:B2", "$ (123.45)", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "B3:B3", "$ -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "C1:C1", "£ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "C2:C2", "-£ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "C3:C3", "£ -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "D1:D1", "€  123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "D2:D2", "€  (123.45)", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "D3:D3", "€  -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "E1:E1", "¥ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "E2:E2", "¥ -123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "E3:E3", "¥ -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "F1:F1", "CHF  123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "F2:F2", "CHF  -123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "F3:F3", "CHF  -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "G1:G1", "₩ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "G2:G2", "-₩ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "G3:G3", "₩ -", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "H1:H1", "£ 123.45", null)]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "H2:H2", "-£ 123.45", "Red")]
        [InlineData("SH151-Custom-Cell-Format-Currency.xlsx", "Sheet1", "H3:H3", "£ -", null)]

        [InlineData("SH152-Custom-Cell-Format.xlsx", "Sheet1", "A1:A1", "1,234,567.0000", null)]
        [InlineData("SH152-Custom-Cell-Format.xlsx", "Sheet1", "B1:B1", "This is the value: abc", null)]

        [InlineData("SH201-Cell-C1-Without-R-Attr.xlsx", "Sheet1", "C1:C1", "3", null)]
        [InlineData("SH202-Cell-C1-D1-Without-R-Attr.xlsx", "Sheet1", "C1:C1", "3", null)]
        [InlineData("SH203-Cell-C1-D1-E1-Without-R-Attr.xlsx", "Sheet1", "C1:C1", "3", null)]
        [InlineData("SH204-Cell-A1-B1-C1-Without-R-Attr.xlsx", "Sheet1", "A1:A1", "1", null)]
        
        public void CF002(string name, string sheetName, string range, string expected, string expectedColor)
        {
            FileInfo sourceXlsx = new FileInfo(Path.Combine(TestUtil.SourceDir.FullName, name));

            var sourceCopiedToDestXlsx = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx")));
            if (!sourceCopiedToDestXlsx.Exists)
                File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);

            var dataTemplateFileNameSuffix = string.Format("-2-Generated-XmlData-{0}.xml", range.Replace(":", ""));
            var dataXmlFi = new FileInfo(Path.Combine(TestUtil.TempDir.FullName, sourceXlsx.Name.Replace(".xlsx", dataTemplateFileNameSuffix)));
            using (SpreadsheetDocument sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, true))
            {
                var rangeXml = SmlDataRetriever.RetrieveRange(sDoc, sheetName, range);
                string displayValue = (string)rangeXml.Descendants("DisplayValue").FirstOrDefault();
                string displayColor = (string)rangeXml.Descendants("DisplayColor").FirstOrDefault();
                Assert.Equal(expected, displayValue);
                Assert.Equal(expectedColor, displayColor);
            }
        }


    }
}

#endif
